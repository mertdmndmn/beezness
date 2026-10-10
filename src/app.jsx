import { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { createClient } from "@supabase/supabase-js";
import ExcelJS from "exceljs";

// ---- Supabase config -------------------------------------------------
// Fill these in with your project's URL and anon (public) key — see
// supabase/schema.sql for the tables + RLS this app expects.
const SUPABASE_URL = "https://biwrbrlsgurqnckzigjg.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJpd3JicmxzZ3VycW5ja3ppZ2pnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODg1MDk2NTcsImV4cCI6MjEwNDA4NTY1N30.gopwHPQWqajYhJT9PQPZKKqAd7-qzaZEqBRT358Y8Dc";

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true },
});

const TABLES = ["locations", "products", "accounts", "batches", "sales", "transfers", "markets", "sale_changes", "expenses"];
const EMPTY_DATA = { locations: [], products: [], accounts: [], batches: [], sales: [], transfers: [], markets: [], sale_changes: [], expenses: [] };
// Newer than the rest of the schema: a project that hasn't run that part of
// schema.sql yet must still load, just without sale history (and without
// the ability to move sales, since that would be a change with no record).
// Same for expenses: no table, no Expenses tab.
const OPTIONAL_TABLES = new Set(["sale_changes", "expenses"]);

const CACHE_KEY = "honey-till-cache-v1";
const OUTBOX_KEY = "honey-till-outbox-v1";
const FAILED_OUTBOX_KEY = "honey-till-outbox-failed-v1";
const LEGACY_KEY = "honey-till-v5"; // last on-device-only schema, kept for one-time import
const MIGRATED_FLAG = "honey-till-migrated-v1";
const DISMISSED_MARKET_KEY = "honey-till-dismissed-market";
// A tip is recorded as a sale line with this stand-in product id, so it
// lands in the account, Who owes and the Z report like any other money,
// but never counts as an item sold or touches stock.
const TIP_PID = "tip";
const unitsOf = (rows) => rows.reduce((s, x) => s + (x.pid === TIP_PID ? 0 : x.qty), 0);
const MY_NAME_KEY = "honey-till-my-name";

const PRODUCT_TYPES = ["Honey", "Candle", "Lip balm", "Other"];
const PAYMENT_METHODS = ["TWINT", "Cash", "Bank transfer"];

const uid = () =>
  typeof crypto !== "undefined" && crypto.randomUUID
    ? crypto.randomUUID()
    : Math.random().toString(36).slice(2) + Date.now().toString(36);

const money = (n) => (Math.round(n * 100) / 100).toFixed(2);
const round2 = (n) => Math.round(n * 100) / 100;
const round5 = (n) => Math.round(n / 0.05) * 0.05;

// Quick cash amounts above the total: nearest 5 (coin round), then the
// Swiss note denominations. Capped at 4 so the row fits on a phone.
function roundCashOptions(total) {
  const steps = [5, 10, 20, 50, 100, 200];
  const seen = new Set();
  const options = [];
  for (const step of steps) {
    const val = Math.ceil(total / step) * step;
    if (val > total + 0.001 && !seen.has(val)) {
      seen.add(val);
      options.push(val);
    }
    if (options.length >= 4) break;
  }
  return options;
}
const dayKey = (ts) => new Date(ts).toDateString();
const timeStr = (ts) => new Date(ts).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
const shortDate = (ts) => new Date(ts).toLocaleDateString([], { day: "2-digit", month: "short" });
const itemsLabel = (lines) => lines.map((s) => `${s.qty}× ${s.name}`).join(", ");

// One entry per checkout (all lines of a ticket share the account they
// were paid to), newest first.
function groupTickets(sales) {
  const byTicket = {};
  sales.forEach((s) => {
    const t = (byTicket[s.ticket] ||= { ticket: s.ticket, ts: s.ts, account: s.account, accountId: s.accountId, method: s.method, note: s.note, market: s.market, lines: [], amount: 0 });
    t.lines.push(s);
    t.amount += s.qty * s.price;
  });
  return Object.values(byTicket).sort((a, b) => b.ts - a.ts);
}
const isoDate = (d) => {
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
};

// ---- Z report styling (ExcelJS) -----------------------------------------
const XLSX_MONEY_FMT = "#,##0.00";
const XLSX_INT_FMT = "0";
const XLSX_HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFBEEDD" } };
const XLSX_SECTION_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2EFE9" } };
const XLSX_BORDER = { style: "thin", color: { argb: "FFE7E2D8" } };

function styleTitleRow(row) {
  row.font = { bold: true, size: 13 };
}
function styleHeaderRow(row) {
  row.font = { bold: true };
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.fill = XLSX_HEADER_FILL;
    cell.border = { bottom: XLSX_BORDER };
  });
}
function styleSectionRow(row) {
  row.font = { bold: true };
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.fill = XLSX_SECTION_FILL;
  });
}
function styleTotalRow(row) {
  row.font = { bold: true };
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.border = { top: XLSX_BORDER };
  });
}

// By type, then alphabetical by name; numeric so "250g" sorts before "1000g". Letter sizes
// ("Candle XS", "Candle M") sort small to big rather than alphabetically.
const LETTER_SIZES = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL"];
const sizeKey = (name) =>
  (name || "").replace(/\b(XXXL|XXL|XL|XXS|XS|S|M|L)\b/g, (m, _s, i, str) =>
    // "1 L" is a volume, not a size.
    /\d\s?$/.test(str.slice(0, i)) ? m : String(LETTER_SIZES.indexOf(m) + 1)
  );
// Type comes first, in PRODUCT_TYPES order, so honey always leads.
const typeRank = (t) => {
  const i = PRODUCT_TYPES.indexOf(t);
  return i === -1 ? PRODUCT_TYPES.length : i;
};
const byName = (a, b) =>
  typeRank(a.type) - typeRank(b.type) ||
  sizeKey(a.name).localeCompare(sizeKey(b.name), undefined, { numeric: true, sensitivity: "base" });

// Products that share a name apart from their size ("Castagno 250 g",
// "Castagno 500 g") are one family. Each family with several sizes gets its
// own row of tiles; one-off products share rows between them. Expects the
// list already sorted by name.
const familyOf = (name) =>
  (name || "")
    .replace(/\d+([.,]\d+)?\s*(kg|g|gr|ml|cl|l|pcs|pz|stk)?\b\.?/gi, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
function tileRows(products) {
  const families = [];
  for (const p of products) {
    const last = families[families.length - 1];
    if (last && familyOf(last[0].name) === familyOf(p.name)) last.push(p);
    else families.push([p]);
  }
  const rows = [];
  for (const fam of families) {
    const last = rows[rows.length - 1];
    if (fam.length === 1 && last && last.loose) last.push(fam[0]);
    else rows.push(Object.assign([...fam], { loose: fam.length === 1 }));
  }
  return rows;
}

function seedProduct(name, type, priceTi, priceLu) {
  return { id: uid(), name, type, price: { "loc-ti": priceTi, "loc-lu": priceLu } };
}
const DEFAULT_LOCATIONS = [
  { id: "loc-ti", name: "Ticino" },
  { id: "loc-lu", name: "Luzern" },
];
const DEFAULT_PRODUCTS = [
  seedProduct("Blossom honey 500 g", "Honey", 16, 18),
  seedProduct("Forest honey 500 g", "Honey", 18, 20),
  seedProduct("Blossom honey 250 g", "Honey", 9, 10),
  seedProduct("Beeswax candle", "Candle", 12, 12),
  seedProduct("Lip balm", "Lip balm", 6, 6),
];
const DEFAULT_ACCOUNTS = [
  { id: uid(), name: "Mert", method: "TWINT", common: false },
  { id: uid(), name: "Cash box", method: "Cash", common: false },
  { id: uid(), name: "Common honey account", method: "Bank transfer", common: true },
];

// ---- local write-through cache + outbox -------------------------------
// Every mutation lands here first so the till keeps working with no
// signal; the outbox is replayed against Supabase whenever we're back
// online. Because ids are generated on the client, replaying an insert
// twice is harmless (upsert), so retries never risk duplicate rows.

function loadCache() {
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}
// Which market (if any) this specific phone chose to leave without ending
// it for everyone else. Cleared implicitly once that market is no longer
// the open one.
function loadDismissedMarket() {
  try {
    return localStorage.getItem(DISMISSED_MARKET_KEY);
  } catch {
    return null;
  }
}
function saveDismissedMarket(id) {
  try {
    localStorage.setItem(DISMISSED_MARKET_KEY, id);
  } catch {}
}
function clearDismissedMarket() {
  try {
    localStorage.removeItem(DISMISSED_MARKET_KEY);
  } catch {}
}
function loadMyName() {
  try {
    return localStorage.getItem(MY_NAME_KEY) || "";
  } catch {
    return "";
  }
}
function saveMyName(name) {
  try {
    localStorage.setItem(MY_NAME_KEY, name);
  } catch {}
}
function saveCache(data) {
  try {
    localStorage.setItem(CACHE_KEY, JSON.stringify(data));
  } catch {}
}
function loadOutbox() {
  try {
    const raw = localStorage.getItem(OUTBOX_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}
function saveOutbox(list) {
  try {
    localStorage.setItem(OUTBOX_KEY, JSON.stringify(list));
  } catch {}
}
function loadFailedOutbox() {
  try {
    const raw = localStorage.getItem(FAILED_OUTBOX_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}
function saveFailedOutbox(list) {
  try {
    localStorage.setItem(FAILED_OUTBOX_KEY, JSON.stringify(list));
  } catch {}
}

// The UI subscribes here to react to outbox/failure changes that happen
// outside React (enqueue, flush, permanent failure) without polling.
const outboxListeners = new Set();
const outboxFailureListeners = new Set();
function notifyOutbox() {
  const snapshot = { pending: loadOutbox().length, failed: loadFailedOutbox() };
  outboxListeners.forEach((fn) => fn(snapshot));
}

function enqueueOps(ops) {
  if (!ops || !ops.length) return;
  const q = loadOutbox();
  q.push(...ops.map((o) => ({ opId: uid(), ...o })));
  saveOutbox(q);
  notifyOutbox();
  flushOutbox();
}

async function applyOp(op) {
  if (op.type === "append") {
    // Plain insert for append-only tables (no update policy, so upsert is
    // refused). A replay after a lost response hits the same client id —
    // that duplicate means it already landed.
    const { error } = await supabase.from(op.table).insert(op.row);
    if (error && error.code !== "23505") throw error;
  } else if (op.type === "insert") {
    const { error } = await supabase.from(op.table).upsert(op.row);
    if (error) throw error;
  } else if (op.type === "update") {
    const { error } = await supabase.from(op.table).update(op.row).eq("id", op.id);
    if (error) throw error;
  } else if (op.type === "delete") {
    const { error } = await supabase.from(op.table).delete().eq("id", op.id);
    if (error) throw error;
  }
}

let flushing = false;
async function flushOutbox() {
  if (flushing) return;
  flushing = true;
  try {
    for (;;) {
      const q = loadOutbox();
      if (!q.length) break;
      const op = q[0];
      try {
        await applyOp(op);
      } catch (err) {
        if (err && err.code) {
          // The server actually answered and rejected this write (bad
          // foreign key, RLS, stale id from a since-replaced project) —
          // retrying changes nothing, and leaving it at the head of the
          // queue would jam every write behind it forever. Park it.
          saveOutbox(loadOutbox().filter((x) => x.opId !== op.opId));
          const failed = loadFailedOutbox();
          failed.push({ ...op, error: err.message || String(err), code: err.code, failedAt: Date.now() });
          saveFailedOutbox(failed);
          notifyOutbox();
          outboxFailureListeners.forEach((fn) => fn(op));
          continue;
        }
        break; // offline or server hiccup — stop, keep order, retry later
      }
      saveOutbox(loadOutbox().filter((x) => x.opId !== op.opId));
      notifyOutbox();
    }
  } finally {
    flushing = false;
  }
}

function mergeById(list, row) {
  const idx = list.findIndex((x) => x.id === row.id);
  if (idx === -1) return [row, ...list];
  const next = list.slice();
  next[idx] = row;
  return next;
}
function removeById(list, id) {
  return list.filter((x) => x.id !== id);
}

async function fetchAll() {
  const results = await Promise.all(TABLES.map((t) => supabase.from(t).select("*")));
  const out = { missing: [] };
  TABLES.forEach((t, i) => {
    const r = results[i];
    if (r.error) {
      if (!OPTIONAL_TABLES.has(t)) throw r.error;
      out.missing.push(t);
    }
    out[t] = r.data || [];
  });
  return out;
}

// One-time carry-over from the old per-device localStorage app: turns
// each product's old mutable `stock` number into an opening batch, since
// stock is never stored directly any more.
async function migrateLegacy(raw) {
  const locations = raw.locations && raw.locations.length ? raw.locations : DEFAULT_LOCATIONS;
  const accounts = (raw.accounts && raw.accounts.length ? raw.accounts : DEFAULT_ACCOUNTS).map((a) => ({
    id: a.id || uid(),
    name: a.name,
    method: a.method || "TWINT",
    common: !!a.common,
  }));
  const idByName = Object.fromEntries(accounts.map((a) => [a.name, a.id]));

  const products = (raw.products || []).map((p) => {
    const priceIsObj = p.price && typeof p.price === "object";
    return {
      id: p.id || uid(),
      name: p.name,
      type: p.type || "Other",
      price: priceIsObj ? { ...p.price } : Object.fromEntries(locations.map((l) => [l.id, Number(p.price) || 0])),
    };
  });

  const batches = [];
  for (const p of raw.products || []) {
    const stockVal =
      p.stock && typeof p.stock === "object"
        ? Object.values(p.stock).reduce((s, v) => s + (Number(v) || 0), 0)
        : Number(p.stock) || 0;
    if (stockVal) {
      batches.push({ id: uid(), pid: p.id || products.find((x) => x.name === p.name)?.id, qty: stockVal, ts: Date.now(), note: "carried over from device storage" });
    }
  }

  const sales = (raw.sales || []).map((s) => ({
    id: s.id || uid(),
    ticket: s.ticket || s.id || uid(),
    ts: s.ts,
    pid: s.pid,
    name: s.name,
    type: s.type,
    price: s.price,
    list: typeof s.list === "number" ? s.list : s.price,
    mode: s.mode || "full",
    qty: s.qty,
    note: s.note || "",
    locId: s.locId || locations[0].id,
    location: s.location || locations[0].name,
    accountId: s.accountId || idByName[s.account] || "",
    account: s.account,
    method: s.method,
  }));

  const transfers = raw.transfers || [];

  const inserts = [
    locations.length && supabase.from("locations").insert(locations),
    products.length && supabase.from("products").insert(products),
    accounts.length && supabase.from("accounts").insert(accounts),
    batches.length && supabase.from("batches").insert(batches),
    sales.length && supabase.from("sales").insert(sales),
    transfers.length && supabase.from("transfers").insert(transfers),
  ].filter(Boolean);
  await Promise.all(inserts);
}

function EditableField({ value, onCommit, mono, cls, style, placeholder, list }) {
  const [text, setText] = useState(String(value));
  useEffect(() => setText(String(value)), [value]);
  return (
    <input
      className={cls}
      style={style}
      list={list}
      placeholder={placeholder}
      value={text}
      inputMode={mono ? "decimal" : undefined}
      onChange={(e) => setText(e.target.value)}
      onBlur={() => {
        if (mono) {
          const n = parseFloat(String(text).replace(",", "."));
          onCommit(isNaN(n) ? 0 : n);
        } else onCommit(text);
      }}
      onKeyDown={(e) => e.key === "Enter" && e.target.blur()}
    />
  );
}

function App() {
  const [authReady, setAuthReady] = useState(false);
  const [session, setSession] = useState(null);
  const [ready, setReady] = useState(false);
  const [data, setData] = useState(() => ({ ...EMPTY_DATA, ...loadCache() }));

  const [tab, setTab] = useState("sell");
  const [activeLoc, setActiveLoc] = useState(() => (loadCache()?.locations || [])[0]?.id || "");
  const [cart, setCart] = useState([]);
  const [note, setNote] = useState("");
  const [pickedProduct, setPickedProduct] = useState(null);
  const [cashAccount, setCashAccount] = useState(null);
  const [cashInput, setCashInput] = useState("");
  const [dismissedMarket, setDismissedMarket] = useState(loadDismissedMarket);
  const [outboxStatus, setOutboxStatus] = useState({ pending: 0, failed: [] });
  const lastStockTapRef = useRef({});
  const [qty, setQty] = useState(1);
  const [editLine, setEditLine] = useState(null); // index of the basket line being re-priced, or null
  const [unitPrice, setUnitPrice] = useState(0);
  const [priceMode, setPriceMode] = useState("full");
  const [receipt, setReceipt] = useState(null);
  const [toast, setToast] = useState(null);
  const [reportDay, setReportDay] = useState(() => dayKey(Date.now()));
  const [reportLoc, setReportLoc] = useState("all");
  const [settlingId, setSettlingId] = useState(null);
  const [payAmount, setPayAmount] = useState("");
  const [payNote, setPayNote] = useState("");
  const [marketDraft, setMarketDraft] = useState(null);
  const [summaryMarket, setSummaryMarket] = useState(null);
  const [confirming, setConfirming] = useState(null); // "today" | "all" | "reset" | "acct:<id>" | null
  // Which optional tables are missing on the server: null until a fetch
  // confirms either way (e.g. offline at boot). Moving a sale needs
  // sale_changes; recording an expense needs expenses.
  const [missingTables, setMissingTables] = useState(null);
  const historyReady = missingTables && !missingTables.includes("sale_changes");
  const expensesReady = missingTables && !missingTables.includes("expenses");
  const [expDraft, setExpDraft] = useState(null); // { paidBy, what, amount, date } while the add form is open
  const [moving, setMoving] = useState(null); // { ticket, toId, by, note } while the move form is open
  const [showAllSales, setShowAllSales] = useState(false);
  const [tipping, setTipping] = useState(null); // { amount } while the market tip form is open
  const [repaying, setRepaying] = useState(null); // { ticket, by } while a market sale's payment is being corrected
  // Pay screen, phone not in market mode while a market is open: did the
  // seller say this sale belongs to it? null = not touched, use the default.
  const [tagMarket, setTagMarket] = useState(null);

  // ---- auth: silent, automatic, no sign-in screen. Everyone who loads the
  // site gets an anonymous session (still "authenticated" for RLS purposes),
  // so this is a shared-link tool, not an access-controlled one.
  useEffect(() => {
    let cancelled = false;
    (async () => {
      const { data: { session } } = await supabase.auth.getSession();
      if (cancelled) return;
      if (session) {
        setSession(session);
        setAuthReady(true);
        return;
      }
      const { data, error } = await supabase.auth.signInAnonymously();
      if (cancelled) return;
      if (!error) setSession(data.session);
      setAuthReady(true);
    })();
    const { data: sub } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
      setAuthReady(true);
    });
    return () => {
      cancelled = true;
      sub.subscription.unsubscribe();
    };
  }, []);

  useEffect(() => {
    if (authReady) setReady(true);
  }, [authReady]);

  useEffect(() => {
    if (data.locations.length && !data.locations.some((l) => l.id === activeLoc)) {
      setActiveLoc(data.locations[0].id);
    }
  }, [data.locations]);

  const applyServer = (server) => {
    const pending = loadOutbox();
    const pendingInsertsByTable = {};
    const pendingDeletesByTable = {};
    pending.forEach((op) => {
      const set = (op.type === "delete" ? pendingDeletesByTable : pendingInsertsByTable);
      (set[op.table] ||= new Set()).add(op.id);
    });
    setData((d) => {
      const next = {};
      TABLES.forEach((table) => {
        // A row still on the server whose delete just hasn't synced yet
        // must not come back to life just because a fetch beat the delete
        // there — otherwise refreshing right after deleting something
        // undoes it.
        const serverList = server[table].filter((s) => !pendingDeletesByTable[table]?.has(s.id));
        const localOnly = (d[table] || []).filter(
          (r) => pendingInsertsByTable[table]?.has(r.id) && !serverList.some((s) => s.id === r.id)
        );
        next[table] = [...serverList, ...localOnly];
      });
      saveCache(next);
      return next;
    });
  };

  // ---- initial sync: fetch, seed, or migrate ----
  useEffect(() => {
    if (!session) return;
    (async () => {
      try {
        const server = await fetchAll();
        const allEmpty = TABLES.every((t) => server[t].length === 0);
        const legacyRaw = localStorage.getItem(LEGACY_KEY);
        if (allEmpty && legacyRaw && !localStorage.getItem(MIGRATED_FLAG)) {
          // One-time carry-over from this specific phone's old per-device
          // data. Deliberately NOT paired with a "seed sample data if the
          // shared database looks empty" fallback any more — that's what
          // silently injected placeholder products into the shared project
          // when a real write just hadn't landed there yet.
          try {
            await migrateLegacy(JSON.parse(legacyRaw));
            localStorage.setItem(MIGRATED_FLAG, "1");
          } catch {}
          applyServer(await fetchAll());
        } else {
          applyServer(server);
        }
        setMissingTables(server.missing);
      } catch {
        // offline at boot — keep showing whatever the cache already has
      }
    })();
  }, [Boolean(session)]);

  // ---- coming back to the app (unlock, switching back from TWINT) pulls a
  // fresh copy, so a market started on another phone a minute ago is here
  // before the next sale — not only once live sync happens to deliver it ----
  useEffect(() => {
    if (!session) return;
    const onVisible = async () => {
      if (document.visibilityState !== "visible") return;
      try {
        const server = await fetchAll();
        applyServer(server);
        setMissingTables(server.missing);
      } catch {}
    };
    document.addEventListener("visibilitychange", onVisible);
    return () => document.removeEventListener("visibilitychange", onVisible);
  }, [Boolean(session)]);

  // ---- outbox retry loop ----
  useEffect(() => {
    if (!session) return;
    flushOutbox();
    const onOnline = () => flushOutbox();
    window.addEventListener("online", onOnline);
    const interval = setInterval(flushOutbox, 20000);
    return () => {
      window.removeEventListener("online", onOnline);
      clearInterval(interval);
    };
  }, [Boolean(session)]);

  // ---- surface outbox state, and drop cache rows whose write was
  // permanently rejected (e.g. an id from a project that no longer exists)
  // instead of leaving a phantom row on screen ----
  useEffect(() => {
    outboxListeners.add(setOutboxStatus);
    setOutboxStatus({ pending: loadOutbox().length, failed: loadFailedOutbox() });
    const onFailure = (op) => {
      if (op.type !== "insert") return;
      setData((d) => {
        const next = { ...d, [op.table]: (d[op.table] || []).filter((r) => r.id !== op.id) };
        saveCache(next);
        return next;
      });
    };
    outboxFailureListeners.add(onFailure);
    return () => {
      outboxListeners.delete(setOutboxStatus);
      outboxFailureListeners.delete(onFailure);
    };
  }, []);

  // ---- pick up a new deploy without waiting for a manual reopen. sw.js's
  // own bytes never change between deploys (only app.js/index.html do,
  // already served network-first), so a service-worker version check would
  // never fire — instead poll app.js's own ETag/Last-Modified directly.
  // Carts aren't persisted, so this never reloads out from under an open
  // sale — it waits until the cart is empty again. ----
  const [updateReady, setUpdateReady] = useState(false);
  useEffect(() => {
    let knownTag = null;
    let cancelled = false;
    const check = async () => {
      try {
        const res = await fetch(`app.js?_cb=${Date.now()}`, { method: "HEAD", cache: "no-store" });
        const tag = res.headers.get("etag") || res.headers.get("last-modified");
        if (!tag || cancelled) return;
        if (knownTag === null) knownTag = tag;
        else if (tag !== knownTag) setUpdateReady(true);
      } catch {}
    };
    check();
    const interval = setInterval(check, 5 * 60 * 1000);
    window.addEventListener("online", check);
    return () => {
      cancelled = true;
      clearInterval(interval);
      window.removeEventListener("online", check);
    };
  }, []);
  useEffect(() => {
    if (updateReady && cart.length === 0) window.location.reload();
  }, [updateReady, cart.length]);

  // ---- realtime: other phones' changes land here without a refresh.
  // Optional tables get their own channel, opened only once the table is
  // known to exist, so a project without them can't break the main one. ----
  const subscribeTables = (name, tables) => {
    const channel = supabase.channel(name);
    tables.forEach((table) => {
      channel.on("postgres_changes", { event: "*", schema: "public", table }, (payload) => {
        setData((d) => {
          const list = d[table];
          const next =
            payload.eventType === "DELETE" ? removeById(list, payload.old.id) : mergeById(list, payload.new);
          const nd = { ...d, [table]: next };
          saveCache(nd);
          return nd;
        });
      });
    });
    return channel;
  };
  useEffect(() => {
    if (!session || !missingTables) return;
    const present = [...OPTIONAL_TABLES].filter((t) => !missingTables.includes(t));
    if (!present.length) return;
    const channel = subscribeTables("honey-till-sync-optional", present).subscribe();
    return () => supabase.removeChannel(channel);
  }, [Boolean(session), String(missingTables)]);
  useEffect(() => {
    if (!session) return;
    const channel = subscribeTables(
      "honey-till-sync",
      TABLES.filter((t) => !OPTIONAL_TABLES.has(t))
    );
    channel.subscribe((status) => {
      if (status === "SUBSCRIBED") flushOutbox();
    });
    return () => supabase.removeChannel(channel);
  }, [Boolean(session)]);

  const save = (partial, ops) => {
    setData((d) => {
      const next = { ...d, ...partial };
      saveCache(next);
      return next;
    });
    enqueueOps(ops);
  };

  // ---- derived state ----
  const stockByPid = useMemo(() => {
    const m = {};
    data.products.forEach((p) => (m[p.id] = 0));
    data.batches.forEach((b) => (m[b.pid] = (m[b.pid] || 0) + Number(b.qty || 0)));
    data.sales.forEach((s) => (m[s.pid] = (m[s.pid] || 0) - Number(s.qty || 0)));
    return m;
  }, [data.products, data.batches, data.sales]);
  const stockOf = (pid) => round2(stockByPid[pid] || 0);

  const locationName = (id) => (data.locations.find((l) => l.id === id) || {}).name || "—";
  const priceOf = (product, locId = activeLoc) => Number((product.price || {})[locId] || 0);
  const inCartQty = (pid) => cart.filter((l) => l.pid === pid).reduce((s, l) => s + l.qty, 0);
  const cartTotal = cart.reduce((s, l) => s + l.qty * l.price, 0);

  const openMarket = data.markets.find((m) => !m.endedAt) || null;
  const activeMarket = openMarket && openMarket.id !== dismissedMarket ? openMarket : null;
  const homeTab = activeMarket ? "market" : "sell";
  // A market is running but this phone isn't in it (left it, or it only
  // just arrived). Sales here get asked whether they belong to it, and
  // default to yes when the phone is set to the market's place.
  const outsideMarket = openMarket && !activeMarket ? openMarket : null;
  const tagToMarket = outsideMarket ? (tagMarket ?? outsideMarket.locId === activeLoc) : false;
  const saleMarket = activeMarket || (tagToMarket ? outsideMarket : null);
  const rejoinMarket = () => {
    clearDismissedMarket();
    setDismissedMarket(null);
    setTab("market");
  };
  const contextLocId = activeMarket ? activeMarket.locId : activeLoc;
  const effectivePrice = (product) => {
    if (activeMarket) {
      const item = activeMarket.items.find((i) => i.pid === product.id);
      if (item) return Number(item.price) || 0;
    }
    return priceOf(product);
  };

  // ---- guard against a market that's stale on this phone: ended, or
  // deleted, elsewhere while it was offline. Markets have the worst signal
  // of anywhere this app runs, so this can't wait for the next full sync —
  // it re-checks on load and whenever connectivity comes back. ----
  useEffect(() => {
    if (!session || !openMarket) return;
    let cancelled = false;
    const verify = async () => {
      const { data: rows, error } = await supabase.from("markets").select("id,endedAt").eq("id", openMarket.id);
      if (cancelled || error) return; // can't confirm — leave it, try again later
      const stillOpen = rows.length && !rows[0].endedAt;
      if (!stillOpen) {
        setData((d) => {
          const next = { ...d, markets: d.markets.filter((m) => m.id !== openMarket.id) };
          saveCache(next);
          return next;
        });
      }
    };
    verify();
    window.addEventListener("online", verify);
    return () => {
      cancelled = true;
      window.removeEventListener("online", verify);
    };
  }, [Boolean(session), openMarket && openMarket.id]);

  const today = useMemo(() => {
    const day = dayKey(Date.now());
    const list = data.sales.filter((s) => dayKey(s.ts) === day && s.locId === activeLoc).sort((a, b) => b.ts - a.ts);
    return {
      list,
      units: unitsOf(list),
      cash: list.reduce((s, x) => s + x.qty * x.price, 0),
    };
  }, [data.sales, activeLoc]);

  const byType = useMemo(() => {
    const m = {};
    [...data.products].sort(byName).forEach((p) => (m[p.type] = m[p.type] || []).push(p));
    return m;
  }, [data.products]);

  const noteSuggestions = useMemo(
    () => [...new Set(data.sales.map((s) => s.note).filter(Boolean))].slice(0, 40),
    [data.sales]
  );

  const saleDays = useMemo(() => [...new Set(data.sales.map((s) => dayKey(s.ts)))], [data.sales]);

  // Markets newest first, for reporting on a whole market however many
  // days it ran. A market's sales are found by marketId, not by date.
  const reportMarkets = useMemo(() => [...data.markets].sort((a, b) => b.startedAt - a.startedAt), [data.markets]);
  const marketSpan = (m) => {
    const from = shortDate(m.startedAt);
    if (!m.endedAt) return `${from} – ongoing`;
    const to = shortDate(m.endedAt);
    return from === to ? from : `${from} – ${to}`;
  };

  const report = useMemo(() => {
    // reportDay holds a dayKey, "market:<id>", or "all" for every sale ever.
    const allTime = reportDay === "all";
    const marketId = reportDay.startsWith("market:") ? reportDay.slice(7) : null;
    const market = marketId ? data.markets.find((m) => m.id === marketId) || { id: marketId, name: "Market" } : null;
    const rows = data.sales
      .filter((s) => (allTime || (market ? s.marketId === market.id : dayKey(s.ts) === reportDay)) && (reportLoc === "all" || s.locId === reportLoc))
      .sort((a, b) => a.ts - b.ts);
    const byDay = {};
    const byAcct = {};
    const byItem = {};
    const byLoc = {};
    const byMarket = {};
    let gifts = 0;
    let discount = 0;
    rows.forEach((s) => {
      const acctKey = `${s.account} · ${s.method}`;
      byAcct[acctKey] = (byAcct[acctKey] || 0) + s.qty * s.price;
      byLoc[s.location] = (byLoc[s.location] || 0) + s.qty * s.price;
      if (s.marketId) byMarket[s.market] = (byMarket[s.market] || 0) + s.qty * s.price;
      const d = new Date(s.ts).toLocaleDateString([], { weekday: "short", day: "numeric", month: "short" });
      byDay[d] = (byDay[d] || 0) + s.qty * s.price;
      const itemKey = `${s.name} · ${s.location}`;
      byItem[itemKey] = byItem[itemKey] || { qty: 0, sum: 0 };
      byItem[itemKey].qty += s.qty;
      byItem[itemKey].sum += s.qty * s.price;
      if (s.mode === "gift") gifts += s.qty;
      else if (s.price < s.list) discount += s.qty * (s.list - s.price);
    });
    return {
      rows,
      market,
      allTime,
      // Sales spread over more than one day get a date next to their time.
      multiDay: allTime || !!market,
      byDay,
      total: rows.reduce((s, x) => s + x.qty * x.price, 0),
      units: unitsOf(rows),
      tickets: new Set(rows.map((s) => s.ticket)).size,
      byMarket,
      gifts,
      discount,
      byAcct,
      byItem,
      byLoc,
    };
  }, [data.sales, data.markets, reportDay, reportLoc]);

  // Every sale lands in exactly one bucket — a person, the common account,
  // or unassigned (its account can't be found, e.g. it was removed in
  // Admin) — so taken === people + common + unassigned and nothing can
  // quietly fall out of the settlement.
  const settleInfo = useMemo(() => {
    const byId = Object.fromEntries(data.accounts.map((a) => [a.id, a]));
    const resolve = (s) => (s.accountId ? byId[s.accountId] : data.accounts.find((a) => a.name === s.account)) || null;
    const weekAgo = Date.now() - 168 * 3600 * 1000;
    const sum = (list) => list.reduce((s, x) => s + x.qty * x.price, 0);
    const salesBy = {};
    const commonSales = [];
    const unassignedSales = [];
    data.sales.forEach((s) => {
      const a = resolve(s);
      if (!a) unassignedSales.push(s);
      else if (a.common) commonSales.push(s);
      else (salesBy[a.id] ||= []).push(s);
    });
    const rows = data.accounts
      .filter((a) => !a.common)
      .map((a) => {
        const theirSales = salesBy[a.id] || [];
        const collected = sum(theirSales);
        const week = sum(theirSales.filter((s) => s.ts >= weekAgo));
        const paid = data.transfers.filter((t) => t.accountId === a.id).reduce((s, x) => s + x.amount, 0);
        return { ...a, collected, week, paid, due: round2(collected - paid) };
      });
    return {
      rows,
      outstanding: rows.reduce((s, x) => s + x.due, 0),
      taken: sum(data.sales),
      heldByPeople: rows.reduce((s, x) => s + x.collected, 0),
      commonAccount: data.accounts.find((a) => a.common) || null,
      common: sum(commonSales),
      commonTickets: groupTickets(commonSales),
      unassigned: sum(unassignedSales),
      unassignedTickets: groupTickets(unassignedSales),
    };
  }, [data.sales, data.transfers, data.accounts]);

  // ---- actions ----
  // Tapping a tile drops one at full price straight into the basket; tapping
  // again adds another. Price changes (half, gift, custom) happen by tapping
  // the line in the basket, which opens the product screen for that line.
  const quickAdd = (product) => {
    if (stockOf(product.id) - inCartQty(product.id) <= 0) return;
    const list = effectivePrice(product);
    const next = [...cart];
    const idx = next.findIndex((l) => l.pid === product.id && l.price === list && l.mode === "full");
    if (idx >= 0) next[idx] = { ...next[idx], qty: next[idx].qty + 1 };
    else next.push({ pid: product.id, name: product.name, type: product.type, price: list, list, mode: "full", qty: 1 });
    setCart(next);
  };

  const changeLineQty = (i, delta) => {
    const line = cart[i];
    if (!line) return;
    if (delta > 0 && stockOf(line.pid) - inCartQty(line.pid) <= 0) return;
    const q = line.qty + delta;
    setCart(q <= 0 ? cart.filter((_, j) => j !== i) : cart.map((l, j) => (j === i ? { ...l, qty: q } : l)));
  };

  const pickProduct = (product, lineIdx = null) => {
    const line = lineIdx === null ? null : cart[lineIdx];
    setPickedProduct(product);
    setEditLine(line ? lineIdx : null);
    setQty(line ? line.qty : 1);
    setUnitPrice(line ? line.price : effectivePrice(product));
    setPriceMode(line ? line.mode : "full");
  };

  const addToCart = (goToPay) => {
    const product = pickedProduct;
    const list = effectivePrice(product);
    const price = priceMode === "gift" ? 0 : Number(unitPrice) || 0;
    const next = editLine === null ? [...cart] : cart.filter((_, j) => j !== editLine);
    setEditLine(null);
    const idx = next.findIndex((l) => l.pid === product.id && l.price === price && l.mode === priceMode);
    if (idx >= 0) next[idx] = { ...next[idx], qty: next[idx].qty + qty };
    else next.push({ pid: product.id, name: product.name, type: product.type, price, list, mode: priceMode, qty });
    setCart(next);
    setPickedProduct(null);
    setTab(goToPay ? "pay" : homeTab);
  };

  const checkout = (account, cashInfo) => {
    const ticket = uid();
    const ts = Date.now();
    const total = cartTotal;
    const lines = cart.map((l) => ({
      id: uid(),
      ticket,
      ts,
      pid: l.pid,
      name: l.name,
      type: l.type,
      price: l.price,
      list: l.list,
      mode: l.mode,
      qty: l.qty,
      note: note.trim(),
      locId: contextLocId,
      location: locationName(contextLocId),
      accountId: account.id,
      account: account.name,
      method: account.method,
      marketId: saleMarket ? saleMarket.id : null,
      market: saleMarket ? saleMarket.name : null,
      cashReceived: cashInfo ? cashInfo.received : null,
      cashChange: cashInfo ? cashInfo.change : null,
    }));
    save(
      { sales: [...lines, ...data.sales] },
      lines.map((row) => ({ table: "sales", type: "insert", id: row.id, row }))
    );
    setReceipt({ total, account: account.name, method: account.method, lines: cart, note: note.trim() });
    setCart([]);
    setTagMarket(null);
    setNote("");
    setTab("done");
    setTimeout(() => {
      setTab((cur) => (cur === "done" ? homeTab : cur));
      setToast({ msg: `CHF ${money(total)} → ${account.name}`, undo: ticket });
    }, 1600);
  };

  // Money between family: a move never overwrites silently. The history
  // row is queued ahead of the sale updates (the outbox replays in order)
  // and carries who, when, and what the sale said before.
  const moveTicket = (ticket, toAccount, changedBy, reason) => {
    const rows = data.sales.filter((s) => s.ticket === ticket);
    const by = changedBy.trim();
    if (!rows.length || !toAccount || !by || historyReady !== true) return;
    const first = rows[0];
    const change = {
      id: uid(),
      ts: Date.now(),
      ticket,
      saleIds: rows.map((r) => r.id),
      items: itemsLabel(rows),
      amount: round2(rows.reduce((s, x) => s + x.qty * x.price, 0)),
      fromAccountId: first.accountId || null,
      fromAccount: first.account || null,
      fromMethod: first.method || null,
      toAccountId: toAccount.id,
      toAccount: toAccount.name,
      toMethod: toAccount.method,
      changedBy: by,
      authUserId: session?.user?.id || null,
      note: reason.trim() || null,
    };
    const patch = { accountId: toAccount.id, account: toAccount.name, method: toAccount.method };
    save(
      {
        sales: data.sales.map((s) => (s.ticket === ticket ? { ...s, ...patch } : s)),
        sale_changes: [change, ...data.sale_changes],
      },
      [
        { table: "sale_changes", type: "append", id: change.id, row: change },
        ...rows.map((r) => ({ table: "sales", type: "update", id: r.id, row: patch })),
      ]
    );
    saveMyName(by);
    setMoving(null);
    setToast({ msg: `CHF ${money(change.amount)} moved to ${toAccount.name}.` });
  };

  // Same matching rule as settleInfo: by id, or by name for rows that
  // predate accountId.
  const accountUsage = (a) => {
    const sales = data.sales.filter((s) => (s.accountId ? s.accountId === a.id : s.account === a.name));
    const transfers = data.transfers.filter((t) => t.accountId === a.id);
    return {
      sales: sales.length,
      amount: sales.reduce((s, x) => s + x.qty * x.price, 0),
      transfers: transfers.length,
      paid: transfers.reduce((s, x) => s + x.amount, 0),
    };
  };
  const removeAccount = (a) => {
    save({ accounts: data.accounts.filter((x) => x.id !== a.id) }, [{ table: "accounts", type: "delete", id: a.id }]);
    setConfirming(null);
  };

  const deleteTicket = (ticket) => {
    const rows = data.sales.filter((s) => s.ticket === ticket);
    if (!rows.length) return;
    save(
      { sales: data.sales.filter((s) => s.ticket !== ticket) },
      rows.map((row) => ({ table: "sales", type: "delete", id: row.id }))
    );
    setToast(null);
  };

  const retryFailedOp = (opId) => {
    const failed = loadFailedOutbox();
    const op = failed.find((x) => x.opId === opId);
    if (!op) return;
    saveFailedOutbox(failed.filter((x) => x.opId !== opId));
    const { error, code, failedAt, ...clean } = op;
    const q = loadOutbox();
    q.push(clean);
    saveOutbox(q);
    notifyOutbox();
    flushOutbox();
  };

  const discardFailedOp = (opId) => {
    saveFailedOutbox(loadFailedOutbox().filter((x) => x.opId !== opId));
    notifyOutbox();
  };

  const retryAllFailedOps = () => {
    const failed = loadFailedOutbox();
    if (!failed.length) return;
    saveFailedOutbox([]);
    const q = loadOutbox();
    q.push(...failed.map(({ error, code, failedAt, ...clean }) => clean));
    saveOutbox(q);
    notifyOutbox();
    flushOutbox();
  };

  const syncFailurePanel = outboxStatus.failed.length > 0 && (
    <div
      style={{ background: "var(--clay-soft)", border: "1px solid var(--clay)", borderRadius: 14, padding: "14px 15px", margin: "0 0 12px" }}
    >
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
        <div style={{ fontSize: 14, lineHeight: 1.5, fontWeight: 600 }}>
          ⚠ {outboxStatus.failed.length} write{outboxStatus.failed.length === 1 ? "" : "s"} rejected by the server — these won't
          retry on their own
        </div>
        <button className="ghost tiny" style={{ flex: "none" }} onClick={retryAllFailedOps}>
          Retry all
        </button>
      </div>
      {outboxStatus.failed.map((op) => (
        <div key={op.opId} style={{ marginTop: 10, fontSize: 13 }}>
          <div>
            <b>{op.type}</b> · {op.table} · {new Date(op.failedAt).toLocaleString()}
          </div>
          <div style={{ opacity: 0.8 }}>{op.error}</div>
          <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
            <button className="ghost" onClick={() => retryFailedOp(op.opId)}>
              Retry
            </button>
            <button className="ghost danger" onClick={() => discardFailedOp(op.opId)}>
              Discard
            </button>
          </div>
        </div>
      ))}
    </div>
  );

  const endMarket = () => {
    const endedAt = Date.now();
    const ended = { ...activeMarket, endedAt };
    save(
      { markets: data.markets.map((m) => (m.id === activeMarket.id ? ended : m)) },
      [{ table: "markets", type: "update", id: activeMarket.id, row: { endedAt } }]
    );
    setSummaryMarket(ended);
    setTab("sell");
  };

  const addTip = (account, amount) => {
    if (!(amount > 0)) return;
    const row = {
      id: uid(),
      ticket: uid(),
      ts: Date.now(),
      pid: TIP_PID,
      name: "Tip",
      type: "Tip",
      price: amount,
      list: amount,
      mode: "full",
      qty: 1,
      note: "",
      locId: contextLocId,
      location: locationName(contextLocId),
      accountId: account.id,
      account: account.name,
      method: account.method,
      marketId: activeMarket ? activeMarket.id : null,
      market: activeMarket ? activeMarket.name : null,
      cashReceived: null,
      cashChange: null,
    };
    save({ sales: [row, ...data.sales] }, [{ table: "sales", type: "insert", id: row.id, row }]);
    setTipping(null);
    setToast({ msg: `Tip CHF ${money(amount)} → ${account.name}`, undo: row.ticket });
  };

  // Sales lists (newest first) keep only the latest few sales open so the
  // buttons below stay in reach; the rest fold away behind a toggle. Lines of
  // one sale stay together.
  const RECENT_SALES = 3;
  const foldSales = (list) => {
    const tickets = [...new Set(list.map((s) => s.ticket))];
    const hidden = Math.max(0, tickets.length - RECENT_SALES);
    if (showAllSales || !hidden) return { shown: list, hidden };
    const recent = new Set(tickets.slice(0, RECENT_SALES));
    return { shown: list.filter((s) => recent.has(s.ticket)), hidden };
  };
  const salesToggle = (hidden) =>
    hidden > 0 && (
      <button className="ghost tiny wide mt8" onClick={() => setShowAllSales(!showAllSales)}>
        {showAllSales ? "▴ Show only the latest" : `▾ Show ${hidden} earlier sale${hidden === 1 ? "" : "s"}`}
      </button>
    );

  const basketPanel = cart.length > 0 && (
    <div className="bag">
      <div className="cap">Basket · tap an item to change its price</div>
      <div style={{ maxHeight: "32vh", overflowY: "auto" }}>
      {cart.map((l, i) => (
        <div className="line" key={i}>
          <span style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0 }}>
            <button className="step" onClick={() => changeLineQty(i, -1)} aria-label="One less">
              −
            </button>
            <b className="num">{l.qty}</b>
            <button className="step" onClick={() => changeLineQty(i, 1)} aria-label="One more">
              +
            </button>
            <button
              className="x"
              style={{ textAlign: "left", fontSize: 14, color: "inherit", padding: 0 }}
              onClick={() => {
                const product = data.products.find((p) => p.id === l.pid);
                if (product) pickProduct(product, i);
              }}
            >
              {l.name}
              {l.mode === "gift" ? " (gift)" : l.price < l.list ? " (reduced)" : ""}
            </button>
          </span>
          <span className="num">{money(l.qty * l.price)}</span>
        </div>
      ))}
      </div>
      <button className="ghost solid wide mt8" style={{ padding: 18, fontSize: 17 }} onClick={() => setTab("pay")}>
        Next · CHF {money(cartTotal)} →
      </button>
    </div>
  );

  const tipPanel = tipping ? (
    <div style={{ background: "var(--ground)", border: "1px solid var(--comb)", borderRadius: 12, padding: "12px 13px", margin: "6px 0 12px" }}>
      <div className="cap mb6">Tip — how much?</div>
      <input
        autoFocus
        inputMode="decimal"
        placeholder="CHF"
        value={tipping.amount}
        onChange={(e) => setTipping({ amount: e.target.value })}
      />
      <div className="cap mt16 mb6">Which account is it going to?</div>
      {data.accounts.map((a) => {
        const amount = round2(parseFloat(String(tipping.amount).replace(",", ".")) || 0);
        return (
          <button key={a.id} className="who" disabled={!(amount > 0)} onClick={() => addTip(a, amount)}>
            <b>{a.name}</b>
            <span>
              {a.method}
              {a.common ? " · common" : ""}
            </span>
          </button>
        );
      })}
      <button className="ghost tiny mt8" onClick={() => setTipping(null)}>
        Cancel
      </button>
    </div>
  ) : (
    <div className="tiles mb10">
      <button className="tile" onClick={() => setTipping({ amount: "" })}>
        <b>Tip</b>
        <div className="pr">any amount</div>
      </button>
    </div>
  );

  // Markets can run over several days: reopen an ended one so the next
  // day's sales land in the same market (and its Z report).
  // "Closed for good" needs the markets.closedAt column (see schema.sql).
  // Rows fetched from the server carry the key, even as null, once it exists.
  const closingReady = data.markets.some((m) => "closedAt" in m);
  const closeMarketForGood = (m) => {
    const closedAt = Date.now();
    const endedAt = m.endedAt || closedAt;
    save(
      { markets: data.markets.map((x) => (x.id === m.id ? { ...x, endedAt, closedAt } : x)) },
      [{ table: "markets", type: "update", id: m.id, row: { endedAt, closedAt } }]
    );
    setToast({ msg: `${m.name} is closed for good.` });
  };

  const reopenMarket = (m) => {
    if (openMarket) return;
    save(
      { markets: data.markets.map((x) => (x.id === m.id ? { ...x, endedAt: null } : x)) },
      [{ table: "markets", type: "update", id: m.id, row: { endedAt: null } }]
    );
    clearDismissedMarket();
    setDismissedMarket(null);
    setMarketDraft(null);
    setTab("market");
    setToast({ msg: `${m.name} is open again.` });
  };

  const recordPayment = (account) => {
    const amt = parseFloat(String(payAmount).replace(",", "."));
    if (isNaN(amt) || amt === 0) return;
    const row = { id: uid(), ts: Date.now(), accountId: account.id, name: account.name, amount: round2(amt), note: payNote.trim() };
    save({ transfers: [row, ...data.transfers] }, [{ table: "transfers", type: "insert", id: row.id, row }]);
    setSettlingId(null);
    setPayAmount("");
    setPayNote("");
    setToast({ msg: `${account.name} settled CHF ${money(amt)}.` });
  };

  // Money someone paid out of their own pocket for the business (a machine,
  // jars, a stand fee). Beezness owes it back until it's marked paid back.
  const addExpense = () => {
    const amt = parseFloat(String(expDraft.amount).replace(",", "."));
    const paidBy = expDraft.paidBy.trim();
    const what = expDraft.what.trim();
    if (isNaN(amt) || amt <= 0 || !paidBy || !what || expensesReady !== true) return;
    const today = new Date();
    const picked = expDraft.date ? new Date(`${expDraft.date}T12:00:00`) : today;
    const ts = picked.toDateString() === today.toDateString() ? Date.now() : picked.getTime();
    const row = { id: uid(), ts, paidBy, what, amount: round2(amt), repaidAt: null };
    save({ expenses: [row, ...data.expenses] }, [{ table: "expenses", type: "insert", id: row.id, row }]);
    saveMyName(paidBy);
    setExpDraft(null);
    setToast({ msg: `Beezness owes ${paidBy} CHF ${money(amt)} for ${what}.` });
  };
  const setExpenseRepaid = (e, repaid) => {
    const repaidAt = repaid ? Date.now() : null;
    save({ expenses: data.expenses.map((x) => (x.id === e.id ? { ...x, repaidAt } : x)) }, [{ table: "expenses", type: "update", id: e.id, row: { repaidAt } }]);
  };
  const expenseInfo = useMemo(() => {
    const byPerson = {};
    data.expenses.forEach((e) => {
      const p = (byPerson[e.paidBy] ||= { name: e.paidBy, spent: 0, owed: 0 });
      p.spent += Number(e.amount) || 0;
      if (!e.repaidAt) p.owed += Number(e.amount) || 0;
    });
    const people = Object.values(byPerson).sort((a, b) => b.owed - a.owed || a.name.localeCompare(b.name));
    return {
      people,
      owed: round2(people.reduce((s, p) => s + p.owed, 0)),
      spent: round2(people.reduce((s, p) => s + p.spent, 0)),
      list: [...data.expenses].sort((a, b) => b.ts - a.ts),
    };
  }, [data.expenses]);
  const payerNames = [...new Set([...data.accounts.filter((a) => !a.common).map((a) => a.name), ...data.expenses.map((e) => e.paidBy)])];

  const clearTodaySales = () => {
    const day = dayKey(Date.now());
    const rows = data.sales.filter((s) => dayKey(s.ts) === day);
    if (!rows.length) return;
    save(
      { sales: data.sales.filter((s) => dayKey(s.ts) !== day) },
      rows.map((row) => ({ table: "sales", type: "delete", id: row.id }))
    );
    setConfirming(null);
    setToast({ msg: `Cleared ${rows.length} sale${rows.length === 1 ? "" : "s"} from today.` });
  };

  const clearSalesAndTransfers = () => {
    const ops = [
      ...data.sales.map((row) => ({ table: "sales", type: "delete", id: row.id })),
      ...data.transfers.map((row) => ({ table: "transfers", type: "delete", id: row.id })),
    ];
    save({ sales: [], transfers: [] }, ops);
    setConfirming(null);
    setToast({ msg: "All sales and transfers cleared." });
  };

  const resetEverything = () => {
    const ops = [
      ...data.sales.map((row) => ({ table: "sales", type: "delete", id: row.id })),
      ...data.transfers.map((row) => ({ table: "transfers", type: "delete", id: row.id })),
      ...data.markets.map((row) => ({ table: "markets", type: "delete", id: row.id })),
      ...data.expenses.map((row) => ({ table: "expenses", type: "delete", id: row.id })),
      ...data.batches.map((row) => ({ table: "batches", type: "delete", id: row.id })),
      ...data.products.map((row) => ({ table: "products", type: "delete", id: row.id })),
      ...data.accounts.map((row) => ({ table: "accounts", type: "delete", id: row.id })),
      ...data.locations.map((row) => ({ table: "locations", type: "delete", id: row.id })),
      ...DEFAULT_LOCATIONS.map((row) => ({ table: "locations", type: "insert", id: row.id, row })),
      ...DEFAULT_PRODUCTS.map((row) => ({ table: "products", type: "insert", id: row.id, row })),
      ...DEFAULT_ACCOUNTS.map((row) => ({ table: "accounts", type: "insert", id: row.id, row })),
    ];
    save(
      {
        locations: DEFAULT_LOCATIONS,
        products: DEFAULT_PRODUCTS,
        accounts: DEFAULT_ACCOUNTS,
        batches: [],
        sales: [],
        transfers: [],
        markets: [],
        expenses: [],
      },
      ops
    );
    setActiveLoc(DEFAULT_LOCATIONS[0].id);
    setConfirming(null);
    setToast({ msg: "Reset to a fresh install." });
  };

  const downloadZReport = async () => {
    const { market, allTime, multiDay } = report;
    // A market or all-time report is "as of now" for stock and settlement;
    // a day report is as of that day.
    const day = multiDay ? new Date() : new Date(reportDay);
    const when = allTime
      ? `All time, as of ${day.toLocaleDateString()}`
      : market
      ? `${market.name}${market.startedAt ? ` (${marketSpan(market)})` : ""}`
      : day.toLocaleDateString();
    const saleTime = (ts) => (multiDay ? `${shortDate(ts)} ${timeStr(ts)}` : timeStr(ts));
    const locs = data.locations;
    const wb = new ExcelJS.Workbook();
    wb.creator = "BeeZness";
    wb.created = new Date();

    // ---- Sales ----
    const salesSheet = wb.addWorksheet("Sales", { views: [{ state: "frozen", ySplit: 3 }] });
    salesSheet.columns = [multiDay ? 14 : 8, 11, 11, 12, 24, 10, 6, 10, 12, 14, 10, 15, 11, 24].map((width) => ({ width }));

    styleTitleRow(salesSheet.addRow(["Z report", when, reportLoc === "all" ? "All places" : locationName(reportLoc)]));
    salesSheet.addRow([]);
    styleHeaderRow(
      salesSheet.addRow(["Time", "Sale no.", "Place", "Market", "Item", "Category", "Qty", "List CHF", "Charged CHF", "Line total CHF", "Price", "Paid to", "Method", "Note"])
    );
    report.rows.forEach((s) => {
      const row = salesSheet.addRow([
        saleTime(s.ts),
        s.ticket,
        s.location,
        s.market || "",
        s.name,
        s.type,
        s.qty,
        s.list,
        s.price,
        round2(s.qty * s.price),
        s.mode === "gift" ? "Gift" : s.price < s.list ? "Reduced" : "Full",
        s.account,
        s.method,
        s.note || "",
      ]);
      row.getCell(7).numFmt = XLSX_INT_FMT;
      [8, 9, 10].forEach((c) => (row.getCell(c).numFmt = XLSX_MONEY_FMT));
    });

    salesSheet.addRow([]);
    const salesTotalRow = salesSheet.addRow(["", "", "", "", "", "", report.units, "", "", round2(report.total), "TOTAL", "", "", ""]);
    salesTotalRow.getCell(7).numFmt = XLSX_INT_FMT;
    salesTotalRow.getCell(10).numFmt = XLSX_MONEY_FMT;
    styleTotalRow(salesTotalRow);
    const giftsRow = salesSheet.addRow(["", "", "", "", "", "", report.gifts, "", "", "", "gifted jars", "", "", ""]);
    giftsRow.getCell(7).numFmt = XLSX_INT_FMT;
    const discountRow = salesSheet.addRow(["", "", "", "", "", "", "", "", "", round2(report.discount), "given as discount", "", "", ""]);
    discountRow.getCell(10).numFmt = XLSX_MONEY_FMT;

    const addMoneySection = (title, entries) => {
      salesSheet.addRow([]);
      styleSectionRow(salesSheet.addRow([title, "CHF"]));
      entries.forEach(([k, v]) => {
        salesSheet.addRow([k, round2(v)]).getCell(2).numFmt = XLSX_MONEY_FMT;
      });
    };
    if (multiDay) addMoneySection("Per day", Object.entries(report.byDay));
    addMoneySection("Per account", Object.entries(report.byAcct));
    addMoneySection("Per place", Object.entries(report.byLoc));
    addMoneySection("Per market", Object.entries(report.byMarket));

    salesSheet.addRow([]);
    styleSectionRow(salesSheet.addRow(["Per item", "Qty", "CHF"]));
    Object.entries(report.byItem).forEach(([k, v]) => {
      const row = salesSheet.addRow([k, v.qty, round2(v.sum)]);
      row.getCell(2).numFmt = XLSX_INT_FMT;
      row.getCell(3).numFmt = XLSX_MONEY_FMT;
    });

    // ---- Stock ----
    const stockSheet = wb.addWorksheet("Stock", { views: [{ state: "frozen", ySplit: 3 }] });
    stockSheet.columns = [24, 10, 11, ...locs.map(() => 16), 18].map((width) => ({ width }));

    styleTitleRow(stockSheet.addRow(["Stock", day.toLocaleDateString()]));
    stockSheet.addRow([]);
    styleHeaderRow(stockSheet.addRow(["Item", "Category", "Units left", ...locs.map((l) => `${l.name} price CHF`), `Value at ${locs[0].name} CHF`]));
    const valueColIdx = 4 + locs.length;
    data.products.forEach((p) => {
      const row = stockSheet.addRow([p.name, p.type, stockOf(p.id), ...locs.map((l) => priceOf(p, l.id)), round2(stockOf(p.id) * priceOf(p, locs[0].id))]);
      row.getCell(3).numFmt = XLSX_INT_FMT;
      locs.forEach((_, i) => (row.getCell(4 + i).numFmt = XLSX_MONEY_FMT));
      row.getCell(valueColIdx).numFmt = XLSX_MONEY_FMT;
    });
    stockSheet.addRow([]);
    const stockTotalRow = stockSheet.addRow([
      "TOTAL",
      "",
      data.products.reduce((s, p) => s + stockOf(p.id), 0),
      ...locs.map(() => ""),
      round2(data.products.reduce((s, p) => s + stockOf(p.id) * priceOf(p, locs[0].id), 0)),
    ]);
    stockTotalRow.getCell(3).numFmt = XLSX_INT_FMT;
    stockTotalRow.getCell(valueColIdx).numFmt = XLSX_MONEY_FMT;
    styleTotalRow(stockTotalRow);

    // ---- Settlement ----
    const settleSheet = wb.addWorksheet("Settlement", { views: [{ state: "frozen", ySplit: 3 }] });
    settleSheet.columns = [20, 16, 21, 19, 16].map((width) => ({ width }));

    styleTitleRow(settleSheet.addRow(["Settlement with the common account", day.toLocaleDateString()]));
    settleSheet.addRow([]);
    styleHeaderRow(settleSheet.addRow(["Person", "Holds money as", "Collected all time CHF", "Paid to common CHF", "Still owes CHF"]));
    settleInfo.rows.forEach((a) => {
      const row = settleSheet.addRow([a.name, a.method, round2(a.collected), round2(a.paid), a.due]);
      [3, 4, 5].forEach((c) => (row.getCell(c).numFmt = XLSX_MONEY_FMT));
    });
    settleSheet.addRow([]);
    const outstandingRow = settleSheet.addRow(["", "", "", "TOTAL OUTSTANDING", round2(settleInfo.outstanding)]);
    outstandingRow.getCell(5).numFmt = XLSX_MONEY_FMT;
    styleTotalRow(outstandingRow);

    settleSheet.addRow([]);
    styleSectionRow(settleSheet.addRow(["Every franc taken, all time"]));
    [
      ["Held by people", settleInfo.heldByPeople],
      ["Paid straight to common", settleInfo.common],
      ["Unassigned", settleInfo.unassigned],
    ].forEach(([label, v]) => (settleSheet.addRow([label, "", round2(v)]).getCell(3).numFmt = XLSX_MONEY_FMT));
    const takenRow = settleSheet.addRow(["Taken", "", round2(settleInfo.taken)]);
    takenRow.getCell(3).numFmt = XLSX_MONEY_FMT;
    styleTotalRow(takenRow);

    if (data.sale_changes.length) {
      settleSheet.addRow([]);
      styleSectionRow(settleSheet.addRow(["Changes to sales"]));
      styleHeaderRow(settleSheet.addRow(["Date", "Changed by", "Amount CHF", "From", "To", "Items", "Note"]));
      [...data.sale_changes]
        .sort((a, b) => a.ts - b.ts)
        .forEach((c) => {
          settleSheet
            .addRow([new Date(c.ts).toLocaleString(), c.changedBy, round2(c.amount), c.fromAccount || "", c.toAccount, c.items || "", c.note || ""])
            .getCell(3).numFmt = XLSX_MONEY_FMT;
        });
    }

    settleSheet.addRow([]);
    styleSectionRow(settleSheet.addRow(["Transfers"]));
    styleHeaderRow(settleSheet.addRow(["Date", "Person", "Amount CHF", "Note"]));
    data.transfers.forEach((t) => {
      settleSheet.addRow([new Date(t.ts).toLocaleDateString(), t.name, round2(t.amount), t.note || ""]).getCell(3).numFmt = XLSX_MONEY_FMT;
    });

    try {
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = allTime
        ? `Z-report-all-time-${isoDate(day)}.xlsx`
        : market ? `Z-report-${market.name.replace(/[^\w-]+/g, "-")}.xlsx` : `Z-report-${isoDate(day)}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setToast({ msg: "Z report saved to your files." });
    } catch {
      setToast({ msg: "Download blocked here — use Copy as CSV." });
    }
  };

  const copyCsv = async () => {
    const lines = ["Time,Sale no.,Place,Market,Item,Qty,List CHF,Charged CHF,Line total CHF,Price,Paid to,Method,Note"];
    report.rows.forEach((s) =>
      lines.push(
        [
          report.multiDay ? `${shortDate(s.ts)} ${timeStr(s.ts)}` : timeStr(s.ts),
          s.ticket,
          s.location,
          s.market || "",
          s.name,
          s.qty,
          money(s.list),
          money(s.price),
          money(s.qty * s.price),
          s.mode === "gift" ? "Gift" : s.price < s.list ? "Reduced" : "Full",
          s.account,
          s.method,
          (s.note || "").replace(/,/g, ";"),
        ].join(",")
      )
    );
    lines.push("", "Person,Collected CHF,Paid CHF,Owes CHF");
    settleInfo.rows.forEach((a) => lines.push([a.name, money(a.collected), money(a.paid), money(a.due)].join(",")));
    lines.push(
      "",
      `Held by people,${money(settleInfo.heldByPeople)}`,
      `Paid straight to common,${money(settleInfo.common)}`,
      `Unassigned,${money(settleInfo.unassigned)}`,
      `Taken,${money(settleInfo.taken)}`
    );
    try {
      await navigator.clipboard.writeText(lines.join("\n"));
      setToast({ msg: "Copied — paste into a sheet." });
    } catch {
      setToast({ msg: "Copy blocked by the browser." });
    }
  };

  const moveForm = (t) => (
    <div style={{ background: "var(--ground)", border: "1px solid var(--comb)", borderRadius: 12, padding: "12px 13px", marginTop: 8 }}>
      {historyReady === true ? (
        <>
          <div className="cap mb6">Move to</div>
          <select value={moving.toId} onChange={(e) => setMoving({ ...moving, toId: e.target.value })} style={{ width: "100%" }}>
            <option value="">Choose an account…</option>
            {data.accounts
              .filter((a) => a.id !== t.accountId)
              .map((a) => (
                <option key={a.id} value={a.id}>
                  {a.name} · {a.method}
                  {a.common ? " · common" : ""}
                </option>
              ))}
          </select>
          <input
            value={moving.by}
            onChange={(e) => setMoving({ ...moving, by: e.target.value })}
            placeholder="Your name (required)"
            style={{ width: "100%", marginTop: 8, fontFamily: "'Barlow',sans-serif" }}
          />
          <input
            value={moving.note}
            onChange={(e) => setMoving({ ...moving, note: e.target.value })}
            placeholder="Why (optional)"
            style={{ width: "100%", marginTop: 8, fontFamily: "'Barlow',sans-serif" }}
          />
          <div className="s" style={{ marginTop: 8 }}>
            Kept in the change history below with your name and the time. What it said before stays on record.
          </div>
        </>
      ) : (
        <div style={{ fontSize: 14, lineHeight: 1.5 }}>
          {historyReady === false
            ? "Sale history isn't set up on the server yet, so moving sales is switched off. Run the sale_changes part of supabase/schema.sql in Supabase, then reopen the app."
            : "Can't reach the server to check that sale history is on, so moving is switched off until it can. Nothing changes without a record."}
        </div>
      )}
      <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
        <button className="ghost" onClick={() => setMoving(null)}>
          Cancel
        </button>
        {historyReady === true && (
          <button
            className="ghost solid"
            disabled={!moving.toId || !moving.by.trim()}
            onClick={() => moveTicket(t.ticket, data.accounts.find((a) => a.id === moving.toId), moving.by, moving.note)}
          >
            Move CHF {money(t.amount)}
          </button>
        )}
      </div>
    </div>
  );

  const ticketRow = (t) => (
    <div key={t.ticket} style={{ borderBottom: "1px solid var(--comb)", padding: "4px 0 10px" }}>
      <div className="row" style={{ borderBottom: 0, paddingBottom: 0 }}>
        <div className="grow">
          <div className="t">{itemsLabel(t.lines)}</div>
          <div className="s">
            {shortDate(t.ts)} {timeStr(t.ts)} · paid to {t.account || "—"}
            {t.market ? ` · ${t.market}` : ""}
            {t.note ? ` · ${t.note}` : ""}
          </div>
        </div>
        <div className="v">{money(t.amount)}</div>
      </div>
      {moving && moving.ticket === t.ticket ? (
        moveForm(t)
      ) : (
        <button className="ghost tiny mt8" onClick={() => setMoving({ ticket: t.ticket, toId: "", by: loadMyName(), note: "" })}>
          Move to…
        </button>
      )}
    </div>
  );

  if (!authReady) return <div className="hl"><div className="empty">Opening BeeZness…</div></div>;
  if (!session) {
    return (
      <div className="hl">
        <div className="empty">
          Couldn't connect. Check your internet connection and reload — if this keeps happening, anonymous sign-ins
          may need to be turned on in Supabase (Authentication → Settings → Anonymous sign-ins).
        </div>
      </div>
    );
  }
  if (!ready) return <div className="hl"><div className="empty">Opening BeeZness…</div></div>;

  if (pickedProduct) {
    const listPrice = effectivePrice(pickedProduct);
    const editingQty = editLine !== null && cart[editLine] ? cart[editLine].qty : 0;
    const available = stockOf(pickedProduct.id) - inCartQty(pickedProduct.id) + editingQty;
    const maxQty = Math.max(1, available);
    const previewPrice = priceMode === "gift" ? 0 : Number(unitPrice) || 0;
    return (
      <div className="hl">
        <button
          className="ghost"
          onClick={() => {
            setPickedProduct(null);
            setEditLine(null);
          }}
        >
          ← Back
        </button>
        <div style={{ marginTop: 18 }}>
          <div className="cap">
            {pickedProduct.type} · {locationName(contextLocId)}
          </div>
          <h1 className="xl">{pickedProduct.name}</h1>
          <div className="num sub">
            {available} in stock · list price CHF {money(listPrice)}
          </div>
        </div>
        <div className="qty">
          <button onClick={() => setQty((q) => Math.max(1, q - 1))} disabled={qty <= 1} aria-label="One less">
            −
          </button>
          <span className="n">{qty}</span>
          <button onClick={() => setQty((q) => Math.min(maxQty, q + 1))} disabled={qty >= maxQty} aria-label="One more">
            +
          </button>
          <div style={{ marginLeft: "auto", textAlign: "right" }}>
            <div className="cap">Subtotal</div>
            <div className="big">
              <small>CHF</small>
              {money(qty * previewPrice)}
            </div>
          </div>
        </div>
        <div className="cap mb6">Price each</div>
        <div className="chips">
          <button
            className="chip"
            data-on={priceMode === "full" ? "1" : "0"}
            onClick={() => {
              setPriceMode("full");
              setUnitPrice(listPrice);
            }}
          >
            Full
          </button>
          <button
            className="chip"
            data-on={priceMode === "half" ? "1" : "0"}
            onClick={() => {
              setPriceMode("half");
              setUnitPrice(round2(listPrice / 2));
            }}
          >
            Half
          </button>
          <button
            className="chip"
            data-on={priceMode === "gift" ? "1" : "0"}
            onClick={() => {
              setPriceMode("gift");
              setUnitPrice(0);
            }}
          >
            Gift
          </button>
          <div className="chip-field">
            <span className="cap">CHF</span>
            <EditableField
              mono
              cls="mini"
              value={previewPrice}
              onCommit={(v) => {
                setUnitPrice(v);
                setPriceMode(v === 0 ? "gift" : v === listPrice ? "full" : "custom");
              }}
            />
          </div>
        </div>
        {editLine !== null ? (
          <button className="ghost solid wide mt16" onClick={() => addToCart(false)}>
            Save · back to basket
          </button>
        ) : (
          <>
            <button className="ghost solid wide mt16" onClick={() => addToCart(true)}>
              Confirm · CHF {money(cartTotal + qty * previewPrice)}
            </button>
            <button className="ghost wide mt8" onClick={() => addToCart(false)}>
              Add another item first
            </button>
          </>
        )}
      </div>
    );
  }

  if (tab === "pay" && !cashAccount) {
    return (
      <div className="hl">
        <button className="ghost" onClick={() => setTab(homeTab)}>
          ← Back
        </button>
        <div style={{ margin: "18px 0 12px" }}>
          <div className="cap">{locationName(contextLocId)}</div>
          <h1 className="xl">CHF {money(cartTotal)}</h1>
        </div>
        {cart.map((l, i) => (
          <div className="row" key={i}>
            <div className="grow">
              <div className="t">
                {l.qty}× {l.name}
              </div>
              <div className="s">
                CHF {money(l.price)} each
                {l.mode === "gift" ? " · gift" : l.price < l.list ? ` · was ${money(l.list)}` : ""}
              </div>
            </div>
            <div className="v">{money(l.qty * l.price)}</div>
          </div>
        ))}
        <div className="cap mt16 mb6">Note — who it went to, optional</div>
        <input
          list="notelist"
          placeholder="e.g. Anna from the choir"
          value={note}
          onChange={(e) => setNote(e.target.value)}
          style={{ fontFamily: "'Barlow',sans-serif" }}
        />
        <datalist id="notelist">
          {noteSuggestions.map((n) => (
            <option key={n} value={n} />
          ))}
        </datalist>
        {outsideMarket && (
          <div style={{ background: "var(--honey-soft)", border: "1px solid var(--honey)", borderRadius: 14, padding: "12px 13px", marginTop: 16 }}>
            <div style={{ fontSize: 14, lineHeight: 1.5, fontWeight: 600 }}>
              {outsideMarket.name} is running at {locationName(outsideMarket.locId)}. Is this sale part of it?
            </div>
            <div className="seg" style={{ margin: "10px 0 0" }}>
              <button data-on={tagToMarket ? "1" : "0"} onClick={() => setTagMarket(true)}>
                Yes, count it in {outsideMarket.name}
              </button>
              <button data-on={tagToMarket ? "0" : "1"} onClick={() => setTagMarket(false)}>
                No
              </button>
            </div>
          </div>
        )}
        <div className="cap mt16 mb6">Which account is it going to?</div>
        {data.accounts.length === 0 && <div className="empty">No accounts yet. Add them in Admin first.</div>}
        {data.accounts.map((a) => (
          <button
            key={a.id}
            className="who"
            onClick={() => {
              if (a.method === "Cash") {
                setCashAccount(a);
                setCashInput("");
              } else {
                checkout(a);
              }
            }}
          >
            <b>{a.name}</b>
            <span>
              {a.method}
              {a.common ? " · common" : ""}
            </span>
          </button>
        ))}
      </div>
    );
  }

  if (cashAccount) {
    const total = cartTotal;
    const roundOptions = roundCashOptions(total);
    const received = parseFloat(String(cashInput).replace(",", "."));
    const hasAmount = cashInput.trim() !== "" && !isNaN(received);
    const rawChange = hasAmount ? received - total : null;
    const sufficient = hasAmount && rawChange >= -0.001;
    const roundedChange = sufficient ? round5(rawChange) : null;
    const wasRounded = sufficient && Math.abs(rawChange - roundedChange) > 0.001;

    const complete = () => {
      if (!sufficient) return;
      checkout(cashAccount, { received: round2(received), change: roundedChange });
      setCashAccount(null);
      setCashInput("");
    };

    return (
      <div className="hl">
        <button
          className="ghost"
          onClick={() => {
            setCashAccount(null);
            setCashInput("");
          }}
        >
          ← Back
        </button>
        <div style={{ marginTop: 18 }}>
          <div className="cap">Cash · {cashAccount.name}</div>
          <div className="big">
            <small>CHF</small>
            {money(total)}
          </div>
        </div>

        <div className="cap mb6 mt16">Amount received</div>
        <div className="chips">
          <button className="chip" data-on={cashInput === money(total) ? "1" : "0"} onClick={() => setCashInput(money(total))}>
            CHF {money(total)}
          </button>
          {roundOptions.map((v) => (
            <button key={v} className="chip" data-on={cashInput === String(v) ? "1" : "0"} onClick={() => setCashInput(String(v))}>
              {v}
            </button>
          ))}
        </div>

        <div className="cap mb6 mt16">Or type an amount</div>
        <input
          inputMode="decimal"
          value={cashInput}
          onChange={(e) => setCashInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && e.target.blur()}
          placeholder={money(total)}
        />

        {hasAmount &&
          (sufficient ? (
            <div style={{ marginTop: 30, textAlign: "center" }}>
              <div className="cap">Change</div>
              <div className="num" style={{ fontSize: 52, fontWeight: 700, color: "var(--sage)", lineHeight: 1, marginTop: 6 }}>
                {money(roundedChange)}
              </div>
              {wasRounded && (
                <div className="cap mt8" style={{ textTransform: "none", letterSpacing: 0 }}>
                  Rounded from CHF {money(rawChange)} — 1 and 2 rappen don't exist
                </div>
              )}
            </div>
          ) : (
            <div style={{ marginTop: 30, textAlign: "center" }}>
              <div className="cap">Still missing</div>
              <div className="num" style={{ fontSize: 40, fontWeight: 700, color: "var(--clay)", lineHeight: 1, marginTop: 6 }}>
                {money(Math.abs(rawChange))}
              </div>
            </div>
          ))}

        <button className="ghost solid wide mt16" disabled={!sufficient} onClick={complete}>
          Done
        </button>
      </div>
    );
  }

  if (tab === "done" && receipt) {
    return (
      <div className="hl">
        <div className="done">
          <div className="seal">
            <svg viewBox="0 0 24 24" fill="none" stroke="#2A2118" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round">
              <path d="M4 12.5 L9.5 18 L20 6.5" />
            </svg>
          </div>
          <h2>{receipt.total === 0 ? "Given" : "Sold"}</h2>
          <div className="num" style={{ fontSize: 24 }}>
            CHF {money(receipt.total)}
          </div>
          <div style={{ fontSize: 15, marginTop: 8 }}>
            {receipt.lines.map((l, i) => (
              <div key={i}>
                {l.qty}× {l.name}
              </div>
            ))}
          </div>
          {receipt.note && <div style={{ fontSize: 15, marginTop: 6 }}>{receipt.note}</div>}
          <div className="cap mt12">
            paid to {receipt.account} · {receipt.method}
          </div>
          <button className="ghost" style={{ marginTop: 26 }} onClick={() => setTab(homeTab)}>
            Next sale
          </button>
        </div>
      </div>
    );
  }

  if (tab === "marketSetup1" && marketDraft) {
    const marketNames = [...new Set(data.markets.map((m) => m.name))];
    const recentEnded = data.markets
      .filter((m) => m.endedAt && !m.closedAt && Date.now() - m.endedAt < 14 * 24 * 3600 * 1000)
      .sort((a, b) => b.endedAt - a.endedAt)
      .slice(0, 5);
    return (
      <div className="hl">
        <button className="ghost" onClick={() => setTab("sell")}>
          ← Back
        </button>
        <div style={{ marginTop: 18 }}>
          <h1 className="xl">Start market day</h1>
        </div>
        {recentEnded.length > 0 && (
          <>
            <div className="cap mb6 mt16">Same market as before? Continue it</div>
            {recentEnded.map((m) => (
              <div key={m.id} style={{ display: "flex", gap: 8, alignItems: "stretch" }}>
                <button className="who" style={{ flex: 1, minWidth: 0 }} onClick={() => reopenMarket(m)}>
                  <b>Continue {m.name}</b>
                  <span>
                    {locationName(m.locId)} · {marketSpan(m)}
                  </span>
                </button>
                {closingReady && (
                  <button className="ghost tiny" style={{ flex: "none", alignSelf: "center" }} onClick={() => closeMarketForGood(m)}>
                    Close for good
                  </button>
                )}
              </div>
            ))}
            <div className="cap mb6 mt16">Or start a new one</div>
          </>
        )}
        <div className="cap mb6 mt16">Market name</div>
        <input
          list="marketnames"
          placeholder="e.g. Sunday market, Lugano"
          value={marketDraft.name}
          onChange={(e) => setMarketDraft({ ...marketDraft, name: e.target.value })}
          style={{ fontFamily: "'Barlow',sans-serif" }}
        />
        <datalist id="marketnames">
          {marketNames.map((n) => (
            <option key={n} value={n} />
          ))}
        </datalist>
        <div className="cap mb6 mt16">Start from which place's prices?</div>
        <div className="seg">
          {data.locations.map((l) => (
            <button
              key={l.id}
              data-on={marketDraft.locId === l.id ? "1" : "0"}
              onClick={() => setMarketDraft({ ...marketDraft, locId: l.id })}
            >
              {l.name}
            </button>
          ))}
        </div>
        <button
          className="ghost solid wide mt16"
          disabled={!marketDraft.name.trim()}
          onClick={() => {
            setMarketDraft((d) => ({
              ...d,
              items: [...data.products].sort(byName).map((p) => ({ pid: p.id, checked: false, price: priceOf(p, d.locId) })),
            }));
            setTab("marketSetup2");
          }}
        >
          Next
        </button>
      </div>
    );
  }

  if (tab === "marketSetup2" && marketDraft) {
    const isEditing = !!activeMarket;
    const confirmSetup = () => {
      const ticked = marketDraft.items.filter((i) => i.checked).map((i) => ({ pid: i.pid, price: Number(i.price) || 0 }));
      if (isEditing) {
        // A typo in the name can be fixed here too. Sales carry the market
        // name (for reports and exports), so they follow along.
        const name = marketDraft.name.trim() || activeMarket.name;
        const renamed = name !== activeMarket.name;
        const renamedSales = renamed ? data.sales.filter((s) => s.marketId === activeMarket.id) : [];
        save(
          {
            markets: data.markets.map((m) => (m.id === activeMarket.id ? { ...m, name, items: ticked } : m)),
            ...(renamed && { sales: data.sales.map((s) => (s.marketId === activeMarket.id ? { ...s, market: name } : s)) }),
          },
          [
            { table: "markets", type: "update", id: activeMarket.id, row: renamed ? { name, items: ticked } : { items: ticked } },
            ...renamedSales.map((r) => ({ table: "sales", type: "update", id: r.id, row: { market: name } })),
          ]
        );
      } else {
        const row = {
          id: uid(),
          name: marketDraft.name.trim() || "Market",
          locId: marketDraft.locId,
          items: ticked,
          startedAt: Date.now(),
          endedAt: null,
        };
        save({ markets: [row, ...data.markets] }, [{ table: "markets", type: "insert", id: row.id, row }]);
      }
      setMarketDraft(null);
      setTab("market");
    };
    return (
      <div className="hl">
        <button
          className="ghost"
          onClick={() => {
            if (isEditing) {
              setMarketDraft(null);
              setTab("market");
            } else {
              setTab("marketSetup1");
            }
          }}
        >
          ← Back
        </button>
        <div style={{ marginTop: 18 }}>
          <h1 className="xl">What are we selling today?</h1>
          <div className="cap sub">{marketDraft.name || locationName(marketDraft.locId)}</div>
        </div>
        {isEditing && (
          <>
            <div className="cap mb6 mt16">Market name</div>
            <input
              value={marketDraft.name}
              onChange={(e) => setMarketDraft({ ...marketDraft, name: e.target.value })}
              style={{ fontFamily: "'Barlow',sans-serif", marginBottom: 12 }}
            />
          </>
        )}
        <div className="empty pt0">Tick what's in the car, and adjust the price for this market if it's different.</div>
        {marketDraft.items.map((item, i) => {
          const product = data.products.find((p) => p.id === item.pid);
          if (!product) return null;
          return (
            <div key={item.pid} className="prod">
              <div className="prod-line">
                <button
                  className="chip"
                  data-on={item.checked ? "1" : "0"}
                  onClick={() =>
                    setMarketDraft((d) => ({
                      ...d,
                      items: d.items.map((x, j) => (j === i ? { ...x, checked: !x.checked } : x)),
                    }))
                  }
                >
                  {item.checked ? "✓ " : ""}
                  {product.name}
                </button>
                <span className="lbl">CHF</span>
                <EditableField
                  mono
                  cls="mini"
                  value={item.price}
                  onCommit={(v) =>
                    setMarketDraft((d) => ({
                      ...d,
                      items: d.items.map((x, j) => (j === i ? { ...x, price: v } : x)),
                    }))
                  }
                />
              </div>
            </div>
          );
        })}
        <button className="ghost solid wide mt14" onClick={confirmSetup}>
          {isEditing ? "Save" : "Start market"}
        </button>
      </div>
    );
  }

  if (summaryMarket) {
    const rows = data.sales.filter((s) => s.marketId === summaryMarket.id);
    const total = rows.reduce((s, x) => s + x.qty * x.price, 0);
    const units = unitsOf(rows);
    const byItem = {};
    const byAcct = {};
    rows.forEach((s) => {
      byItem[s.name] = byItem[s.name] || { qty: 0, sum: 0 };
      byItem[s.name].qty += s.qty;
      byItem[s.name].sum += s.qty * s.price;
      byAcct[s.account] = (byAcct[s.account] || 0) + s.qty * s.price;
    });
    return (
      <div className="hl">
        <div style={{ marginTop: 18 }}>
          <div className="cap">Market day done</div>
          <h1 className="xl">{summaryMarket.name}</h1>
        </div>
        <div className="stat">
          <div>
            <div className="cap">Takings</div>
            <div className="n">{money(total)}</div>
          </div>
          <div>
            <div className="cap">Jars sold</div>
            <div className="n">{units}</div>
          </div>
        </div>
        <hr className="rule" />
        <div className="cap">Per product</div>
        {Object.keys(byItem).length === 0 && <div className="empty">Nothing sold.</div>}
        {Object.entries(byItem).map(([k, v]) => (
          <div className="row" key={k}>
            <div className="grow">
              <div className="t">{k}</div>
              <div className="s">{v.qty} sold</div>
            </div>
            <div className="v">{money(v.sum)}</div>
          </div>
        ))}
        {Object.keys(byAcct).length > 0 && (
          <>
            <div className="cap mt16">Per account</div>
            {Object.entries(byAcct).map(([k, v]) => (
              <div className="row" key={k}>
                <div className="grow t">{k}</div>
                <div className="v">{money(v)}</div>
              </div>
            ))}
          </>
        )}
        <button className="ghost solid wide mt16" onClick={() => setSummaryMarket(null)}>
          Back to BeeZness · continue another day
        </button>
        {closingReady && (
          <button
            className="ghost wide mt8"
            onClick={() => {
              closeMarketForGood(summaryMarket);
              setSummaryMarket(null);
            }}
          >
            Last day — close {summaryMarket.name} for good
          </button>
        )}
      </div>
    );
  }

  if (activeMarket) {
    const marketSales = data.sales.filter((s) => s.marketId === activeMarket.id).sort((a, b) => b.ts - a.ts);
    const marketUnits = unitsOf(marketSales);
    const marketCash = marketSales.reduce((s, x) => s + x.qty * x.price, 0);
    const marketProducts = activeMarket.items.map((item) => data.products.find((p) => p.id === item.pid)).filter(Boolean).sort(byName);
    return (
      <div className="hl">
        <div className="top">
          <div>
            <h1>{activeMarket.name}</h1>
            <div className="cap">
              Market day
              {" · "}
              <button
                className="x"
                style={{ fontSize: 13, padding: 0, verticalAlign: "baseline" }}
                onClick={() => window.location.reload()}
                aria-label="Refresh app"
              >
                ⟳ refresh
              </button>
            </div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div className="big">
              <small>CHF</small>
              {money(marketCash)}
            </div>
            <div className="cap" style={{ marginTop: 4 }}>
              {marketUnits} sold
            </div>
          </div>
        </div>

        {syncFailurePanel}

        {marketProducts.length === 0 && (
          <div className="empty">No products picked for today — use "Edit name or today's table" below.</div>
        )}
        {tileRows(marketProducts).map((row) => (
          <div className="tiles mb10" key={row[0].id}>
            {row.map((p) => {
              const available = stockOf(p.id) - inCartQty(p.id);
              return (
                <button key={p.id} className="tile" disabled={available <= 0} onClick={() => quickAdd(p)}>
                  {inCartQty(p.id) > 0 && <span className="inbag">{inCartQty(p.id)}</span>}
                  <b>{p.name}</b>
                  <div className="pr">CHF {money(effectivePrice(p))}</div>
                  <div className="st" data-low={available <= 3 ? "1" : "0"}>
                    {available > 0 ? `${available} left` : "sold out"}
                  </div>
                </button>
              );
            })}
          </div>
        ))}

        {tipPanel}

        {basketPanel}

        <hr className="rule" />
        <div className="cap">Today's sales</div>
        {marketSales.length === 0 && <div className="empty">Nothing sold yet.</div>}
        {foldSales(marketSales).shown.map((s) => (
          <div className="row" key={s.id}>
            <div className="grow">
              <div className="t">
                {s.pid === TIP_PID ? "Tip" : `${s.qty}× ${s.name}`}
                {s.mode === "gift" ? " · gift" : s.price < s.list ? " · reduced" : ""}
              </div>
              <div className="s">
                {timeStr(s.ts)} · {s.account}
                {s.note ? ` · ${s.note}` : ""}
              </div>
              {repaying && repaying.ticket === s.ticket && (
                <div style={{ background: "var(--ground)", border: "1px solid var(--comb)", borderRadius: 12, padding: "10px 11px", marginTop: 8 }}>
                  {historyReady === true ? (
                    <>
                      <div className="cap mb6">Actually paid with</div>
                      {!loadMyName() && (
                        <input
                          value={repaying.by}
                          onChange={(e) => setRepaying({ ...repaying, by: e.target.value })}
                          placeholder="Your name (asked once)"
                          style={{ width: "100%", marginBottom: 8, fontFamily: "'Barlow',sans-serif" }}
                        />
                      )}
                      {data.accounts
                        .filter((a) => a.id !== s.accountId)
                        .map((a) => (
                          <button
                            key={a.id}
                            className="who"
                            disabled={!repaying.by.trim()}
                            onClick={() => {
                              moveTicket(s.ticket, a, repaying.by, "Wrong payment picked at the market");
                              setRepaying(null);
                            }}
                          >
                            <b>{a.name}</b>
                            <span>
                              {a.method}
                              {a.common ? " · common" : ""}
                            </span>
                          </button>
                        ))}
                    </>
                  ) : (
                    <div style={{ fontSize: 14, lineHeight: 1.5 }}>
                      Can't change payments until sale history is reachable on the server — nothing changes without a record.
                    </div>
                  )}
                  <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
                    <button className="ghost tiny" onClick={() => setRepaying(null)}>
                      Cancel
                    </button>
                    {repaying.confirmDelete ? (
                      <button
                        className="ghost tiny danger solid"
                        style={{ marginLeft: "auto" }}
                        onClick={() => {
                          deleteTicket(s.ticket);
                          setRepaying(null);
                        }}
                      >
                        Yes, delete this sale
                      </button>
                    ) : (
                      <button className="ghost tiny danger" style={{ marginLeft: "auto" }} onClick={() => setRepaying({ ...repaying, confirmDelete: true })}>
                        Delete sale…
                      </button>
                    )}
                  </div>
                </div>
              )}
            </div>
            <div className="v">{money(s.qty * s.price)}</div>
            <button
              className="x"
              onClick={() => setRepaying(repaying && repaying.ticket === s.ticket ? null : { ticket: s.ticket, by: loadMyName() })}
              aria-label="Change payment"
            >
              ✎
            </button>
          </div>
        ))}
        {salesToggle(foldSales(marketSales).hidden)}

        <hr className="rule" />
        <button
          className="ghost wide"
          onClick={() => {
            setMarketDraft({
              name: activeMarket.name,
              locId: activeMarket.locId,
              items: [...data.products].sort(byName).map((p) => {
                const found = activeMarket.items.find((i) => i.pid === p.id);
                return { pid: p.id, checked: !!found, price: found ? found.price : priceOf(p, activeMarket.locId) };
              }),
            });
            setTab("marketSetup2");
          }}
        >
          Edit name or today's table
        </button>
        <button className="ghost wide mt8" onClick={endMarket}>
          End market day
        </button>
        <button
          className="ghost wide mt8"
          onClick={() => {
            saveDismissedMarket(activeMarket.id);
            setDismissedMarket(activeMarket.id);
            setTab("sell");
          }}
        >
          Leave market mode
        </button>
      </div>
    );
  }

  return (
    <div className="hl">
      <div className="top">
        <div>
          <h1>BeeZness</h1>
          <div className="cap">
            {new Date().toLocaleDateString([], { weekday: "long", day: "numeric", month: "long" })}
            {" · "}
            <button
              className="x"
              style={{ fontSize: 13, padding: 0, verticalAlign: "baseline" }}
              onClick={() => window.location.reload()}
              aria-label="Refresh app"
            >
              ⟳ refresh
            </button>
          </div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div className="big">
            <small>CHF</small>
            {money(today.cash)}
          </div>
          <div className="cap" style={{ marginTop: 4 }}>
            {today.units} sold · {locationName(activeLoc)}
          </div>
        </div>
      </div>

      {syncFailurePanel}

      {outsideMarket && (
        <div style={{ background: "var(--honey-soft)", border: "1px solid var(--honey)", borderRadius: 14, padding: "14px 15px", margin: "0 0 12px" }}>
          <div style={{ fontSize: 15, lineHeight: 1.45, fontWeight: 600 }}>
            {outsideMarket.name} is running at {locationName(outsideMarket.locId)}
          </div>
          <div style={{ fontSize: 13, color: "var(--soft)", marginTop: 2 }}>
            Started {timeStr(outsideMarket.startedAt)} · selling here uses normal prices
          </div>
          <button className="ghost solid wide mt8" onClick={rejoinMarket}>
            Sell at {outsideMarket.name}
          </button>
        </div>
      )}

      <div className="seg">
        {data.locations.map((l) => (
          <button
            key={l.id}
            data-on={activeLoc === l.id ? "1" : "0"}
            onClick={() => {
              if (cart.length && l.id !== activeLoc) {
                setToast({ msg: "Finish or clear the open sale before switching place." });
                return;
              }
              setActiveLoc(l.id);
            }}
          >
            {l.name}
          </button>
        ))}
      </div>

      <div className="tabs">
        {[
          ["sell", "Sell"],
          ["report", "Z report"],
          ["settle", "Who owes"],
          ["expenses", "Expenses"],
          ["admin", "Admin"],
        ].map(([id, label]) => (
          <button key={id} data-on={tab === id ? "1" : "0"} onClick={() => setTab(id)}>
            {label}
          </button>
        ))}
      </div>

      {tab === "sell" && (
        <>
          <button
            className="ghost solid wide"
            style={{ marginBottom: 16 }}
            onClick={() => {
              setMarketDraft({ name: "", locId: activeLoc || (data.locations[0] || {}).id || "", items: [] });
              setTab("marketSetup1");
            }}
          >
            Start market day
          </button>

          {data.products.length === 0 && (
            <div className="empty">Nothing to sell yet. Open Admin and add your honey, candles and balms.</div>
          )}
          {Object.keys(byType).map((type) => (
            <div key={type} style={{ marginBottom: 18 }}>
              <div className="cap" style={{ marginBottom: 9 }}>
                {type}
              </div>
              {tileRows(byType[type]).map((row) => (
                <div className="tiles mb10" key={row[0].id}>
                  {row.map((p) => {
                    const available = stockOf(p.id) - inCartQty(p.id);
                    return (
                      <button key={p.id} className="tile" disabled={available <= 0} onClick={() => quickAdd(p)}>
                        {inCartQty(p.id) > 0 && <span className="inbag">{inCartQty(p.id)}</span>}
                        <b>{p.name}</b>
                        <div className="pr">CHF {money(priceOf(p))}</div>
                        <div className="st" data-low={available <= 3 ? "1" : "0"}>
                          {available > 0 ? `${available} left` : "sold out"}
                        </div>
                      </button>
                    );
                  })}
                </div>
              ))}
            </div>
          ))}

          {tipPanel}

          {basketPanel}

          <hr className="rule" />
          <div className="cap">Today at {locationName(activeLoc)}</div>
          {today.list.length === 0 && <div className="empty">Nothing sold here yet today.</div>}
          {foldSales(today.list).shown.map((s) => (
            <div className="row" key={s.id}>
              <div className="grow">
                <div className="t">
                  {s.pid === TIP_PID ? "Tip" : `${s.qty}× ${s.name}`}
                  {s.mode === "gift" ? " · gift" : s.price < s.list ? " · reduced" : ""}
                </div>
                <div className="s">
                  {timeStr(s.ts)} · {s.account}
                  {s.note ? ` · ${s.note}` : ""}
                </div>
              </div>
              <div className="v">{money(s.qty * s.price)}</div>
              <button className="x" onClick={() => deleteTicket(s.ticket)} aria-label="Delete this sale">
                ×
              </button>
            </div>
          ))}
          {salesToggle(foldSales(today.list).hidden)}
        </>
      )}

      {tab === "report" && (
        <>
          <div className="two">
            <div>
              <div className="cap mb6">Day or market</div>
              <select value={reportDay} onChange={(e) => setReportDay(e.target.value)}>
                <option value="all">All time</option>
                <optgroup label="Days">
                  {[...new Set([dayKey(Date.now()), ...saleDays])].map((d) => (
                    <option key={d} value={d}>
                      {new Date(d).toLocaleDateString([], { weekday: "short", day: "numeric", month: "long" })}
                    </option>
                  ))}
                </optgroup>
                {reportMarkets.length > 0 && (
                  <optgroup label="Markets">
                    {reportMarkets.map((m) => (
                      <option key={m.id} value={`market:${m.id}`}>
                        {m.name} · {marketSpan(m)}
                      </option>
                    ))}
                  </optgroup>
                )}
              </select>
            </div>
            <div>
              <div className="cap mb6">Where</div>
              <select value={reportLoc} onChange={(e) => setReportLoc(e.target.value)}>
                <option value="all">Everywhere</option>
                {data.locations.map((l) => (
                  <option key={l.id} value={l.id}>
                    {l.name}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <div className="stat">
            <div>
              <div className="cap">Takings</div>
              <div className="n">{money(report.total)}</div>
            </div>
            <div>
              <div className="cap">Items</div>
              <div className="n">{report.units}</div>
            </div>
            <div>
              <div className="cap">Sales</div>
              <div className="n">{report.tickets}</div>
            </div>
          </div>

          {(report.gifts > 0 || report.discount > 0) && (
            <div className="cap mt12">
              {report.gifts > 0 && `${report.gifts} given away`}
              {report.gifts > 0 && report.discount > 0 && " · "}
              {report.discount > 0 && `CHF ${money(report.discount)} off list`}
            </div>
          )}

          <hr className="rule" />
          <div className="cap">Per account</div>
          {Object.keys(report.byAcct).length === 0 && <div className="empty">{report.allTime ? "Nothing sold yet." : report.market ? "Nothing sold at this market." : "Nothing on this day."}</div>}
          {Object.entries(report.byAcct).map(([k, v]) => (
            <div className="row" key={k}>
              <div className="grow t">{k}</div>
              <div className="v">{money(v)}</div>
            </div>
          ))}

          {report.allTime && Object.keys(report.byMarket).length > 0 && (
            <>
              <div className="cap mt16">Per market</div>
              {Object.entries(report.byMarket).map(([k, v]) => (
                <div className="row" key={k}>
                  <div className="grow t">{k}</div>
                  <div className="v">{money(v)}</div>
                </div>
              ))}
            </>
          )}

          {report.market && Object.keys(report.byDay).length > 0 && (
            <>
              <div className="cap mt16">Per day</div>
              {Object.entries(report.byDay).map(([k, v]) => (
                <div className="row" key={k}>
                  <div className="grow t">{k}</div>
                  <div className="v">{money(v)}</div>
                </div>
              ))}
            </>
          )}

          {Object.keys(report.byLoc).length > 1 && (
            <>
              <div className="cap mt16">Per place</div>
              {Object.entries(report.byLoc).map(([k, v]) => (
                <div className="row" key={k}>
                  <div className="grow t">{k}</div>
                  <div className="v">{money(v)}</div>
                </div>
              ))}
            </>
          )}

          <div className="cap mt16">Per item</div>
          {Object.entries(report.byItem).map(([k, v]) => (
            <div className="row" key={k}>
              <div className="grow">
                <div className="t">{k}</div>
                <div className="s">{v.qty} sold</div>
              </div>
              <div className="v">{money(v.sum)}</div>
            </div>
          ))}

          <hr className="rule" />
          <div className="cap">Stock on hand</div>
          <div className="row">
            <div className="grow t">Items in stock</div>
            <div className="v">{data.products.reduce((s, p) => s + Math.max(0, stockOf(p.id)), 0)}</div>
          </div>
          {data.locations.map((l) => (
            <div className="row" key={l.id}>
              <div className="grow t">Worth at {l.name} prices</div>
              <div className="v">{money(data.products.reduce((s, p) => s + Math.max(0, stockOf(p.id)) * priceOf(p, l.id), 0))}</div>
            </div>
          ))}

          <hr className="rule" />
          <button className="ghost solid wide" onClick={downloadZReport}>
            Download Z report (Excel)
          </button>
          <button className="ghost wide mt8" onClick={copyCsv}>
            Copy as CSV instead
          </button>
          <div className="empty">
            Three sheets: {report.allTime ? "every sale ever" : report.market ? "the market's sales" : "the day's sales"} with notes, the stock you have left, and who still owes the common account.
          </div>
        </>
      )}

      {tab === "settle" && (
        <>
          <div className="cap">Owed to the common account</div>
          <div className="big" style={{ fontSize: 34, margin: "6px 0 4px" }}>
            <small>CHF</small>
            {money(settleInfo.outstanding)}
          </div>
          <div className="empty pt0">
            Every sale sticks to the person whose account took the money. Pay in, tap the amount, and the balance clears.
          </div>
          <div className="settle">
            <div className="cap mb6">Every franc taken, all time</div>
            {[
              ["Held by people", settleInfo.heldByPeople],
              ["Paid straight to common", settleInfo.common],
              ...(settleInfo.unassigned ? [["Unassigned", settleInfo.unassigned]] : []),
            ].map(([label, v]) => (
              <div key={label} style={{ display: "flex", justifyContent: "space-between", fontSize: 14, padding: "3px 0", color: label === "Unassigned" ? "var(--clay)" : undefined }}>
                <span>{label}</span>
                <span className="num">{money(v)}</span>
              </div>
            ))}
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 14, fontWeight: 600, padding: "6px 0 0", marginTop: 4, borderTop: "1px solid var(--comb)" }}>
              <span>Taken</span>
              <span className="num">{money(settleInfo.taken)}</span>
            </div>
          </div>

          {settleInfo.unassignedTickets.length > 0 && (
            <div className="settle" style={{ borderColor: "var(--clay)", background: "var(--clay-soft)" }}>
              <div className="settle-head">
                <div className="grow">
                  <div className="t" style={{ fontSize: 17, fontWeight: 600 }}>
                    Unassigned
                  </div>
                  <div className="s">The account these were paid to no longer exists. Attach each one to whoever holds the money.</div>
                </div>
                <div className="num" style={{ fontSize: 20, fontWeight: 600, color: "var(--clay)" }}>
                  {money(settleInfo.unassigned)}
                </div>
              </div>
              {settleInfo.unassignedTickets.map(ticketRow)}
            </div>
          )}

          {settleInfo.rows.length === 0 && (
            <div className="empty">Mark one account as the common one in Admin, and the rest become people who collect for it.</div>
          )}
          {settleInfo.rows.map((a) => (
            <div className="settle" key={a.id}>
              <div className="settle-head">
                <div className="grow">
                  <div className="t" style={{ fontSize: 17, fontWeight: 600 }}>
                    {a.name}
                  </div>
                  <div className="s">
                    {a.method} · CHF {money(a.week)} in the last 7 days
                  </div>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div className="cap">Owes</div>
                  <div className="num" style={{ fontSize: 20, fontWeight: 600, color: a.due > 0 ? "var(--clay)" : "var(--sage)" }}>
                    {money(a.due)}
                  </div>
                </div>
              </div>
              <div className="s" style={{ marginTop: 5 }}>
                collected {money(a.collected)} · paid in {money(a.paid)}
              </div>
              {settlingId === a.id ? (
                <div className="payrow">
                  <input
                    inputMode="decimal"
                    value={payAmount}
                    onChange={(e) => setPayAmount(e.target.value)}
                    placeholder={money(a.due)}
                    style={{ width: 100 }}
                  />
                  <input
                    value={payNote}
                    onChange={(e) => setPayNote(e.target.value)}
                    placeholder="note"
                    style={{ flex: 1, fontFamily: "'Barlow',sans-serif" }}
                  />
                  <button className="ghost solid tiny" onClick={() => recordPayment(a)}>
                    Save
                  </button>
                  <button
                    className="ghost tiny"
                    onClick={() => {
                      setSettlingId(null);
                      setPayAmount("");
                      setPayNote("");
                    }}
                  >
                    ×
                  </button>
                </div>
              ) : (
                <button
                  className="ghost wide mt8"
                  onClick={() => {
                    setSettlingId(a.id);
                    setPayAmount(a.due > 0 ? money(a.due) : "");
                  }}
                >
                  Record payment to common
                </button>
              )}
            </div>
          ))}

          <div className="settle">
            <div className="settle-head">
              <div className="grow">
                <div className="t" style={{ fontSize: 17, fontWeight: 600 }}>
                  Paid straight to common
                </div>
                <div className="s">
                  {settleInfo.commonAccount
                    ? `${settleInfo.commonAccount.name} · ${settleInfo.commonAccount.method} · already where it belongs, owed by nobody`
                    : "No account is marked common yet — pick one in Admin"}
                </div>
              </div>
              <div className="num" style={{ fontSize: 20, fontWeight: 600 }}>
                {money(settleInfo.common)}
              </div>
            </div>
            {settleInfo.commonTickets.length === 0 ? (
              <div className="s" style={{ marginTop: 5 }}>
                Nothing yet. Sales paid directly into the common account show up here.
              </div>
            ) : (
              settleInfo.commonTickets.map(ticketRow)
            )}
          </div>

          {data.sale_changes.length > 0 && (
            <>
              <hr className="rule" />
              <div className="cap">Changes to sales</div>
              {[...data.sale_changes]
                .sort((a, b) => b.ts - a.ts)
                .map((c) => (
                  <div className="row" key={c.id}>
                    <div className="grow">
                      <div className="t">
                        {c.changedBy} moved {c.items} from {c.fromAccount || "an unknown account"} to {c.toAccount}
                      </div>
                      <div className="s">
                        {shortDate(c.ts)} {timeStr(c.ts)}
                        {c.note ? ` · ${c.note}` : ""}
                      </div>
                    </div>
                    <div className="v">{money(c.amount)}</div>
                  </div>
                ))}
            </>
          )}

          {data.transfers.length > 0 && (
            <>
              <hr className="rule" />
              <div className="cap">Paid in so far</div>
              {data.transfers.slice(0, 20).map((t) => (
                <div className="row" key={t.id}>
                  <div className="grow">
                    <div className="t">{t.name}</div>
                    <div className="s">
                      {shortDate(t.ts)}
                      {t.note ? ` · ${t.note}` : ""}
                    </div>
                  </div>
                  <div className="v">{money(t.amount)}</div>
                  <button
                    className="x"
                    onClick={() => save({ transfers: data.transfers.filter((x) => x.id !== t.id) }, [{ table: "transfers", type: "delete", id: t.id }])}
                    aria-label="Remove transfer"
                  >
                    ×
                  </button>
                </div>
              ))}
            </>
          )}
        </>
      )}

      {tab === "expenses" && (
        <>
          <div className="cap">Beezness owes back</div>
          <div className="big" style={{ fontSize: 34, margin: "6px 0 4px" }}>
            <small>CHF</small>
            {money(expenseInfo.owed)}
          </div>
          <div className="empty pt0">
            Paid for something out of your own pocket? Write it here, and it stays owed to you until it's marked paid back.
          </div>

          {expensesReady === false ? (
            <div className="settle" style={{ borderColor: "var(--clay)", background: "var(--clay-soft)" }}>
              <div className="s">
                Expenses aren't switched on yet. Run the expenses part of supabase/schema.sql in the Supabase SQL editor, then refresh.
              </div>
            </div>
          ) : expDraft ? (
            <div className="settle">
              <div className="cap mb6">New expense</div>
              <input
                list="expense-payers"
                value={expDraft.paidBy}
                onChange={(e) => setExpDraft({ ...expDraft, paidBy: e.target.value })}
                placeholder="Who paid"
                style={{ width: "100%", marginBottom: 7, fontFamily: "'Barlow',sans-serif" }}
              />
              <datalist id="expense-payers">
                {payerNames.map((n) => (
                  <option key={n} value={n} />
                ))}
              </datalist>
              <input
                value={expDraft.what}
                onChange={(e) => setExpDraft({ ...expDraft, what: e.target.value })}
                placeholder="What for, e.g. honey extractor"
                style={{ width: "100%", marginBottom: 7, fontFamily: "'Barlow',sans-serif" }}
              />
              <div className="payrow" style={{ marginTop: 0 }}>
                <input
                  inputMode="decimal"
                  value={expDraft.amount}
                  onChange={(e) => setExpDraft({ ...expDraft, amount: e.target.value })}
                  placeholder="CHF"
                  style={{ width: 100 }}
                />
                <input
                  type="date"
                  value={expDraft.date}
                  onChange={(e) => setExpDraft({ ...expDraft, date: e.target.value })}
                  style={{ flex: 1, minWidth: 0, fontFamily: "'Barlow',sans-serif" }}
                />
              </div>
              <div className="payrow">
                <button
                  className="ghost solid wide"
                  disabled={expensesReady !== true}
                  onClick={addExpense}
                >
                  {expensesReady === true ? "Save expense" : "Waiting for connection…"}
                </button>
                <button className="ghost tiny" onClick={() => setExpDraft(null)}>
                  ×
                </button>
              </div>
            </div>
          ) : (
            <button
              className="ghost solid wide mb6"
              onClick={() => {
                const d = new Date();
                const date = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
                setExpDraft({ paidBy: loadMyName(), what: "", amount: "", date });
              }}
            >
              Add an expense
            </button>
          )}

          {expenseInfo.people.map((p) => (
            <div className="settle" key={p.name}>
              <div className="settle-head">
                <div className="grow">
                  <div className="t" style={{ fontSize: 17, fontWeight: 600 }}>
                    {p.name}
                  </div>
                  <div className="s">spent {money(p.spent)} · paid back {money(p.spent - p.owed)}</div>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div className="cap">Owed</div>
                  <div className="num" style={{ fontSize: 20, fontWeight: 600, color: p.owed > 0 ? "var(--clay)" : "var(--sage)" }}>
                    {money(p.owed)}
                  </div>
                </div>
              </div>
            </div>
          ))}

          {expenseInfo.list.length > 0 && (
            <>
              <hr className="rule" />
              <div className="cap">Every expense</div>
              {expenseInfo.list.map((e) => (
                <div className="row" key={e.id} style={{ opacity: e.repaidAt ? 0.6 : 1 }}>
                  <div className="grow">
                    <div className="t">{e.what}</div>
                    <div className="s">
                      {e.paidBy} · {shortDate(e.ts)}
                      {e.repaidAt ? ` · paid back ${shortDate(e.repaidAt)}` : ""}
                    </div>
                  </div>
                  <div className="v">{money(e.amount)}</div>
                  <button className="ghost tiny" onClick={() => setExpenseRepaid(e, !e.repaidAt)}>
                    {e.repaidAt ? "Undo" : "Paid back"}
                  </button>
                  <button
                    className="x"
                    onClick={() => save({ expenses: data.expenses.filter((x) => x.id !== e.id) }, [{ table: "expenses", type: "delete", id: e.id }])}
                    aria-label="Remove expense"
                  >
                    ×
                  </button>
                </div>
              ))}
            </>
          )}
        </>
      )}

      {tab === "admin" && (
        <>
          <button className="ghost wide mb6" onClick={() => window.location.reload()}>
            Refresh app
          </button>
          <h2>Products</h2>
          <div className="empty pt0">One stock for everything you own. Each place has its own price.</div>
          {data.products.map((p) => (
            <div key={p.id} className="prod">
              <div className="prod-head">
                <EditableField
                  value={p.name}
                  cls="name"
                  onCommit={(v) => save({ products: data.products.map((x) => (x.id === p.id ? { ...x, name: v } : x)) }, [{ table: "products", type: "update", id: p.id, row: { name: v } }])}
                />
                <button
                  className="x"
                  onClick={() =>
                    save(
                      { products: data.products.filter((x) => x.id !== p.id), batches: data.batches.filter((b) => b.pid !== p.id) },
                      [{ table: "products", type: "delete", id: p.id }]
                    )
                  }
                  aria-label="Remove product"
                >
                  ×
                </button>
              </div>
              <div className="prod-line">
                <select
                  value={p.type}
                  onChange={(e) =>
                    save({ products: data.products.map((x) => (x.id === p.id ? { ...x, type: e.target.value } : x)) }, [
                      { table: "products", type: "update", id: p.id, row: { type: e.target.value } },
                    ])
                  }
                >
                  {PRODUCT_TYPES.map((t) => (
                    <option key={t}>{t}</option>
                  ))}
                </select>
                <span className="lbl">Jars</span>
                <EditableField
                  mono
                  cls="mini"
                  value={stockOf(p.id)}
                  onCommit={(v) => {
                    const delta = round2(Number(v) - stockOf(p.id));
                    if (!delta) return;
                    const row = { id: uid(), pid: p.id, qty: delta, ts: Date.now(), note: "adjustment" };
                    save({ batches: [row, ...data.batches] }, [{ table: "batches", type: "insert", id: row.id, row }]);
                  }}
                />
                {[12, 24].map((n) => (
                  <button
                    key={n}
                    className="ghost tiny"
                    onClick={() => {
                      // Guards against a burst of rapid taps on the same
                      // button creating several duplicate stock batches —
                      // there's no visible per-tap feedback here, so an
                      // impatient double/triple tap is easy to trigger.
                      const tapKey = p.id + ":" + n;
                      const now = Date.now();
                      if (now - (lastStockTapRef.current[tapKey] || 0) < 800) return;
                      lastStockTapRef.current[tapKey] = now;
                      const row = { id: uid(), pid: p.id, qty: n, ts: now, note: null };
                      save({ batches: [row, ...data.batches] }, [{ table: "batches", type: "insert", id: row.id, row }]);
                    }}
                  >
                    +{n}
                  </button>
                ))}
              </div>
              <div className="prices">
                {data.locations.map((l) => (
                  <div className="pricebox" key={l.id}>
                    <span className="lbl">{l.name}</span>
                    <EditableField
                      mono
                      cls="mini"
                      value={priceOf(p, l.id)}
                      onCommit={(v) => {
                        const price = { ...p.price, [l.id]: v };
                        save({ products: data.products.map((x) => (x.id === p.id ? { ...x, price } : x)) }, [
                          { table: "products", type: "update", id: p.id, row: { price } },
                        ]);
                      }}
                    />
                  </div>
                ))}
              </div>
            </div>
          ))}
          <button
            className="ghost solid mt14"
            onClick={() => {
              const row = { id: uid(), name: "New product", type: "Honey", price: Object.fromEntries(data.locations.map((l) => [l.id, 0])) };
              save({ products: [...data.products, row] }, [{ table: "products", type: "insert", id: row.id, row }]);
            }}
          >
            Add product
          </button>

          <hr className="rule" />
          <h2>Places</h2>
          <div className="empty pt0">Where you sell, each with its own price list.</div>
          {data.locations.map((l) => (
            <div className="line-edit" key={l.id}>
              <EditableField
                value={l.name}
                cls="name"
                onCommit={(v) =>
                  save({ locations: data.locations.map((x) => (x.id === l.id ? { ...x, name: v } : x)) }, [
                    { table: "locations", type: "update", id: l.id, row: { name: v } },
                  ])
                }
              />
              {data.locations.length > 1 && (
                <button
                  className="x"
                  onClick={() => {
                    const remaining = data.locations.filter((x) => x.id !== l.id);
                    setActiveLoc(remaining[0].id);
                    save({ locations: remaining }, [{ table: "locations", type: "delete", id: l.id }]);
                  }}
                  aria-label="Remove place"
                >
                  ×
                </button>
              )}
            </div>
          ))}
          <button
            className="ghost solid mt14"
            onClick={() => {
              const id = uid();
              const newLoc = { id, name: "New place" };
              const updatedProducts = data.products.map((p) => ({ ...p, price: { ...p.price, [id]: 0 } }));
              save({ locations: [...data.locations, newLoc], products: updatedProducts }, [
                { table: "locations", type: "insert", id, row: newLoc },
                ...updatedProducts.map((p) => ({ table: "products", type: "update", id: p.id, row: { price: p.price } })),
              ]);
            }}
          >
            Add place
          </button>

          <hr className="rule" />
          <h2>Accounts</h2>
          <div className="empty pt0">
            Whose TWINT or cash box the money lands in. Tick the one that is the common honey account — everyone else's takings count as owed to it.
          </div>
          {data.accounts.map((a) => (
            <div className="acct" key={a.id}>
              <div className="line-edit" style={{ border: 0, padding: 0 }}>
                <EditableField
                  value={a.name}
                  cls="name"
                  onCommit={(v) =>
                    save({ accounts: data.accounts.map((x) => (x.id === a.id ? { ...x, name: v } : x)) }, [
                      { table: "accounts", type: "update", id: a.id, row: { name: v } },
                    ])
                  }
                />
                <select
                  value={a.method}
                  onChange={(e) =>
                    save({ accounts: data.accounts.map((x) => (x.id === a.id ? { ...x, method: e.target.value } : x)) }, [
                      { table: "accounts", type: "update", id: a.id, row: { method: e.target.value } },
                    ])
                  }
                  style={{ width: 138 }}
                >
                  {PAYMENT_METHODS.map((m) => (
                    <option key={m}>{m}</option>
                  ))}
                </select>
                <button
                  className="x"
                  onClick={() => {
                    const key = `acct:${a.id}`;
                    if (accountUsage(a).sales || accountUsage(a).transfers) setConfirming(confirming === key ? null : key);
                    else removeAccount(a);
                  }}
                  aria-label="Remove account"
                >
                  ×
                </button>
              </div>
              {confirming === `acct:${a.id}` &&
                (() => {
                  const u = accountUsage(a);
                  return (
                    <div
                      style={{ background: "var(--clay-soft)", border: "1px solid var(--clay)", borderRadius: 14, padding: "14px 15px", marginTop: 8 }}
                    >
                      <div style={{ fontSize: 14, lineHeight: 1.5 }}>
                        {a.name} has {u.sales} sale{u.sales === 1 ? "" : "s"} (CHF {money(u.amount)})
                        {u.transfers ? ` and ${u.transfers} payment${u.transfers === 1 ? "" : "s"} into common (CHF ${money(u.paid)})` : ""}.
                        Removing the account moves those sales to Unassigned on the Who owes tab until they're attached to someone else.
                        {u.transfers ? " Its payments into common will no longer count towards anyone." : ""}
                      </div>
                      <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                        <button className="ghost" onClick={() => setConfirming(null)}>
                          Cancel
                        </button>
                        <button className="ghost danger solid" onClick={() => removeAccount(a)}>
                          Remove {a.name}
                        </button>
                      </div>
                    </div>
                  );
                })()}
              <button
                className="chip mt8"
                data-on={a.common ? "1" : "0"}
                onClick={() => {
                  const updated = data.accounts.map((x) => ({ ...x, common: x.id === a.id ? !a.common : false }));
                  save(
                    { accounts: updated },
                    updated.map((x) => ({ table: "accounts", type: "update", id: x.id, row: { common: x.common } }))
                  );
                }}
              >
                {a.common ? "✓ Common honey account" : "Make this the common account"}
              </button>
            </div>
          ))}
          <button
            className="ghost solid mt14"
            onClick={() => {
              const row = { id: uid(), name: "New person", method: "TWINT", common: false };
              save({ accounts: [...data.accounts, row] }, [{ table: "accounts", type: "insert", id: row.id, row }]);
            }}
          >
            Add account
          </button>

          <hr className="rule" />
          <h2>Danger zone</h2>
          {(() => {
            const todayCount = data.sales.filter((s) => dayKey(s.ts) === dayKey(Date.now())).length;
            const allCount = data.sales.length + data.transfers.length;
            const resetCount =
              data.locations.length +
              data.products.length +
              data.accounts.length +
              data.batches.length +
              data.sales.length +
              data.transfers.length +
              data.markets.length;
            return (
              <>
                <button
                  className="ghost danger wide mt8"
                  disabled={todayCount === 0}
                  onClick={() => setConfirming(confirming === "today" ? null : "today")}
                >
                  Clear today's sales
                </button>
                {confirming === "today" && (
                  <div
                    style={{ background: "var(--clay-soft)", border: "1px solid var(--clay)", borderRadius: 14, padding: "14px 15px", marginTop: 8 }}
                  >
                    <div style={{ fontSize: 14, lineHeight: 1.5 }}>
                      Deletes {todayCount} sale{todayCount === 1 ? "" : "s"} from today and returns those jars to stock.
                    </div>
                    <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                      <button className="ghost" onClick={() => setConfirming(null)}>
                        Cancel
                      </button>
                      <button className="ghost danger solid" onClick={clearTodaySales}>
                        Clear today's sales
                      </button>
                    </div>
                  </div>
                )}

                <button
                  className="ghost danger wide mt8"
                  disabled={allCount === 0}
                  onClick={() => setConfirming(confirming === "all" ? null : "all")}
                >
                  Clear all sales and transfers
                </button>
                {confirming === "all" && (
                  <div
                    style={{ background: "var(--clay-soft)", border: "1px solid var(--clay)", borderRadius: 14, padding: "14px 15px", marginTop: 8 }}
                  >
                    <div style={{ fontSize: 14, lineHeight: 1.5 }}>
                      Deletes {data.sales.length} sale{data.sales.length === 1 ? "" : "s"} and {data.transfers.length} transfer
                      {data.transfers.length === 1 ? "" : "s"}. Products, prices, places and accounts stay untouched, and every
                      sold jar returns to stock. This cannot be undone.
                    </div>
                    <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                      <button className="ghost" onClick={() => setConfirming(null)}>
                        Cancel
                      </button>
                      <button className="ghost danger solid" onClick={clearSalesAndTransfers}>
                        Clear sales &amp; transfers
                      </button>
                    </div>
                  </div>
                )}

                <button
                  className="ghost danger wide mt8"
                  onClick={() => setConfirming(confirming === "reset" ? null : "reset")}
                >
                  Reset everything
                </button>
                {confirming === "reset" && (
                  <div
                    style={{ background: "var(--clay-soft)", border: "1px solid var(--clay)", borderRadius: 14, padding: "14px 15px", marginTop: 8 }}
                  >
                    <div style={{ fontSize: 14, lineHeight: 1.5 }}>
                      Deletes everything — {resetCount} record{resetCount === 1 ? "" : "s"} across places, products, accounts,
                      stock, sales, transfers and markets — and reinstalls the starting sample data. This cannot be undone.
                    </div>
                    <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                      <button className="ghost" onClick={() => setConfirming(null)}>
                        Cancel
                      </button>
                      <button className="ghost danger solid" onClick={resetEverything}>
                        Reset everything
                      </button>
                    </div>
                  </div>
                )}
              </>
            );
          })()}
        </>
      )}

      {toast && (
        <div className="toast">
          <span>{toast.msg}</span>
          {toast.undo ? (
            <button onClick={() => deleteTicket(toast.undo)}>Undo</button>
          ) : (
            <button onClick={() => setToast(null)}>OK</button>
          )}
        </div>
      )}
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
