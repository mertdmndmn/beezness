import { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

// ---- Supabase config -------------------------------------------------
// Fill these in with your project's URL and anon (public) key — see
// supabase/schema.sql for the tables + RLS this app expects.
const SUPABASE_URL = "https://rtzdqjftxbxlrcgderbk.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJ0emRxamZ0eGJ4bHJjZ2RlcmJrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODgwOTQ4NTUsImV4cCI6MjEwMzY3MDg1NX0.BHUqT8yyfcERFI_4O_rx7ETMTgVCNcZ72PtLmBwpgSk";

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true },
});

const TABLES = ["locations", "products", "accounts", "batches", "sales", "transfers", "markets"];
const EMPTY_DATA = { locations: [], products: [], accounts: [], batches: [], sales: [], transfers: [], markets: [] };

const CACHE_KEY = "honey-till-cache-v1";
const OUTBOX_KEY = "honey-till-outbox-v1";
const LEGACY_KEY = "honey-till-v5"; // last on-device-only schema, kept for one-time import
const MIGRATED_FLAG = "honey-till-migrated-v1";
const SEEDED_FLAG = "honey-till-seeded-v1";

const PRODUCT_TYPES = ["Honey", "Candle", "Lip balm", "Other"];
const PAYMENT_METHODS = ["TWINT", "Cash", "Bank transfer"];

const uid = () =>
  typeof crypto !== "undefined" && crypto.randomUUID
    ? crypto.randomUUID()
    : Math.random().toString(36).slice(2) + Date.now().toString(36);

const money = (n) => (Math.round(n * 100) / 100).toFixed(2);
const round2 = (n) => Math.round(n * 100) / 100;
const dayKey = (ts) => new Date(ts).toDateString();
const timeStr = (ts) => new Date(ts).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
const shortDate = (ts) => new Date(ts).toLocaleDateString([], { day: "2-digit", month: "short" });
const isoDate = (d) => {
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
};

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
function enqueueOps(ops) {
  if (!ops || !ops.length) return;
  const q = loadOutbox();
  q.push(...ops.map((o) => ({ opId: uid(), ...o })));
  saveOutbox(q);
  flushOutbox();
}

async function applyOp(op) {
  if (op.type === "insert") {
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
      } catch {
        break; // offline or server hiccup — stop, keep order, retry later
      }
      saveOutbox(loadOutbox().filter((x) => x.opId !== op.opId));
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
  results.forEach((r) => {
    if (r.error) throw r.error;
  });
  const out = {};
  TABLES.forEach((t, i) => (out[t] = results[i].data || []));
  return out;
}

async function seedDefaults() {
  await supabase.from("locations").insert(DEFAULT_LOCATIONS);
  await supabase.from("products").insert(DEFAULT_PRODUCTS);
  await supabase.from("accounts").insert(DEFAULT_ACCOUNTS);
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

function SignIn() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [mode, setMode] = useState("signin"); // "signin" | "signup"
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async () => {
    if (!email.trim() || !password || busy) return;
    setBusy(true);
    setErr("");
    if (mode === "signup") {
      const { data, error } = await supabase.auth.signUp({ email: email.trim(), password });
      setBusy(false);
      if (error) setErr(error.message);
      else if (!data.session) {
        setErr(
          'Account created, but it still needs email confirmation. In Supabase: Authentication → Providers → Email → turn off "Confirm email", then try signing in again — or check your inbox for a confirmation link.'
        );
      }
    } else {
      const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
      setBusy(false);
      if (error) setErr(error.message);
    }
  };

  return (
    <div className="hl">
      <div className="top">
        <div>
          <h1>BeeZness</h1>
          <div className="cap">{mode === "signup" ? "Create an account" : "Sign in to open the till"}</div>
        </div>
      </div>
      <div className="cap mb6 mt16">Email</div>
      <input
        type="email"
        value={email}
        onChange={(e) => setEmail(e.target.value)}
        placeholder="you@example.com"
        style={{ fontFamily: "'Barlow',sans-serif" }}
      />
      <div className="cap mb6 mt16">Password</div>
      <input
        type="password"
        value={password}
        onChange={(e) => setPassword(e.target.value)}
        onKeyDown={(e) => e.key === "Enter" && submit()}
        placeholder="••••••••"
      />
      {err && (
        <div className="cap mt12" style={{ color: "var(--clay)" }}>
          {err}
        </div>
      )}
      <button className="ghost solid wide mt16" onClick={submit} disabled={busy}>
        {busy ? "Please wait…" : mode === "signup" ? "Create account" : "Sign in"}
      </button>
      <button
        className="ghost wide mt8"
        onClick={() => {
          setMode(mode === "signup" ? "signin" : "signup");
          setErr("");
        }}
      >
        {mode === "signup" ? "Already have an account? Sign in" : "New here? Create an account"}
      </button>
    </div>
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
  const [qty, setQty] = useState(1);
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

  // ---- auth ----
  useEffect(() => {
    let cancelled = false;
    supabase.auth.getSession().then(({ data: { session } }) => {
      if (cancelled) return;
      setSession(session);
      setAuthReady(true);
    });
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
    const pendingByTable = {};
    pending.forEach((op) => (pendingByTable[op.table] ||= new Set()).add(op.id));
    setData((d) => {
      const next = {};
      TABLES.forEach((table) => {
        const serverList = server[table];
        const localOnly = (d[table] || []).filter(
          (r) => pendingByTable[table]?.has(r.id) && !serverList.some((s) => s.id === r.id)
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
        if (allEmpty) {
          const legacyRaw = localStorage.getItem(LEGACY_KEY);
          if (legacyRaw && !localStorage.getItem(MIGRATED_FLAG)) {
            try {
              await migrateLegacy(JSON.parse(legacyRaw));
              localStorage.setItem(MIGRATED_FLAG, "1");
            } catch {}
          } else if (!localStorage.getItem(SEEDED_FLAG)) {
            await seedDefaults();
            localStorage.setItem(SEEDED_FLAG, "1");
          }
          applyServer(await fetchAll());
        } else {
          applyServer(server);
        }
      } catch {
        // offline at boot — keep showing whatever the cache already has
      }
    })();
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

  // ---- realtime: other phones' changes land here without a refresh ----
  useEffect(() => {
    if (!session) return;
    const channel = supabase.channel("honey-till-sync");
    TABLES.forEach((table) => {
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

  const activeMarket = data.markets.find((m) => !m.endedAt) || null;
  const homeTab = activeMarket ? "market" : "sell";
  const contextLocId = activeMarket ? activeMarket.locId : activeLoc;
  const effectivePrice = (product) => {
    if (activeMarket) {
      const item = activeMarket.items.find((i) => i.pid === product.id);
      if (item) return Number(item.price) || 0;
    }
    return priceOf(product);
  };

  const today = useMemo(() => {
    const day = dayKey(Date.now());
    const list = data.sales.filter((s) => dayKey(s.ts) === day && s.locId === activeLoc);
    return {
      list,
      units: list.reduce((s, x) => s + x.qty, 0),
      cash: list.reduce((s, x) => s + x.qty * x.price, 0),
    };
  }, [data.sales, activeLoc]);

  const byType = useMemo(() => {
    const m = {};
    data.products.forEach((p) => (m[p.type] = m[p.type] || []).push(p));
    return m;
  }, [data.products]);

  const noteSuggestions = useMemo(
    () => [...new Set(data.sales.map((s) => s.note).filter(Boolean))].slice(0, 40),
    [data.sales]
  );

  const saleDays = useMemo(() => [...new Set(data.sales.map((s) => dayKey(s.ts)))], [data.sales]);

  const report = useMemo(() => {
    const rows = data.sales
      .filter((s) => dayKey(s.ts) === reportDay && (reportLoc === "all" || s.locId === reportLoc))
      .sort((a, b) => a.ts - b.ts);
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
      const itemKey = `${s.name} · ${s.location}`;
      byItem[itemKey] = byItem[itemKey] || { qty: 0, sum: 0 };
      byItem[itemKey].qty += s.qty;
      byItem[itemKey].sum += s.qty * s.price;
      if (s.mode === "gift") gifts += s.qty;
      else if (s.price < s.list) discount += s.qty * (s.list - s.price);
    });
    return {
      rows,
      total: rows.reduce((s, x) => s + x.qty * x.price, 0),
      units: rows.reduce((s, x) => s + x.qty, 0),
      tickets: new Set(rows.map((s) => s.ticket)).size,
      byMarket,
      gifts,
      discount,
      byAcct,
      byItem,
      byLoc,
    };
  }, [data.sales, reportDay, reportLoc]);

  const settleInfo = useMemo(() => {
    const people = data.accounts.filter((a) => !a.common);
    const weekAgo = Date.now() - 168 * 3600 * 1000;
    const rows = people.map((a) => {
      const theirSales = data.sales.filter((s) => (s.accountId ? s.accountId === a.id : s.account === a.name));
      const collected = theirSales.reduce((s, x) => s + x.qty * x.price, 0);
      const week = theirSales.filter((s) => s.ts >= weekAgo).reduce((s, x) => s + x.qty * x.price, 0);
      const paid = data.transfers.filter((t) => t.accountId === a.id).reduce((s, x) => s + x.amount, 0);
      return { ...a, collected, week, paid, due: round2(collected - paid) };
    });
    return { rows, outstanding: rows.reduce((s, x) => s + x.due, 0) };
  }, [data.sales, data.transfers, data.accounts]);

  // ---- actions ----
  const pickProduct = (product) => {
    setPickedProduct(product);
    setQty(1);
    setUnitPrice(effectivePrice(product));
    setPriceMode("full");
  };

  const addToCart = (goToPay) => {
    const product = pickedProduct;
    const list = effectivePrice(product);
    const price = priceMode === "gift" ? 0 : Number(unitPrice) || 0;
    const next = [...cart];
    const idx = next.findIndex((l) => l.pid === product.id && l.price === price && l.mode === priceMode);
    if (idx >= 0) next[idx] = { ...next[idx], qty: next[idx].qty + qty };
    else next.push({ pid: product.id, name: product.name, type: product.type, price, list, mode: priceMode, qty });
    setCart(next);
    setPickedProduct(null);
    setTab(goToPay ? "pay" : homeTab);
  };

  const checkout = (account) => {
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
      marketId: activeMarket ? activeMarket.id : null,
      market: activeMarket ? activeMarket.name : null,
    }));
    save(
      { sales: [...lines, ...data.sales] },
      lines.map((row) => ({ table: "sales", type: "insert", id: row.id, row }))
    );
    setReceipt({ total, account: account.name, method: account.method, lines: cart, note: note.trim() });
    setCart([]);
    setNote("");
    setTab("done");
    setTimeout(() => {
      setTab((cur) => (cur === "done" ? homeTab : cur));
      setToast({ msg: `CHF ${money(total)} → ${account.name}`, undo: ticket });
    }, 1600);
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

  const downloadZReport = () => {
    const day = new Date(reportDay);
    const locs = data.locations;
    const salesRows = [
      ["Z report", day.toLocaleDateString(), reportLoc === "all" ? "All places" : locationName(reportLoc)],
      [],
      ["Time", "Sale no.", "Place", "Market", "Item", "Category", "Qty", "List CHF", "Charged CHF", "Line total CHF", "Price", "Paid to", "Method", "Note"],
      ...report.rows.map((s) => [
        timeStr(s.ts),
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
      ]),
      [],
      ["", "", "", "", "", "", report.units, "", "", round2(report.total), "TOTAL", "", "", ""],
      ["", "", "", "", "", "", report.gifts, "", "", "", "gifted jars", "", "", ""],
      ["", "", "", "", "", "", "", "", "", round2(report.discount), "given as discount", "", "", ""],
      [],
      ["Per account", "CHF"],
      ...Object.entries(report.byAcct).map(([k, v]) => [k, round2(v)]),
      [],
      ["Per place", "CHF"],
      ...Object.entries(report.byLoc).map(([k, v]) => [k, round2(v)]),
      [],
      ["Per market", "CHF"],
      ...Object.entries(report.byMarket).map(([k, v]) => [k, round2(v)]),
      [],
      ["Per item", "Qty", "CHF"],
      ...Object.entries(report.byItem).map(([k, v]) => [k, v.qty, round2(v.sum)]),
    ];
    const salesSheet = XLSX.utils.aoa_to_sheet(salesRows);
    salesSheet["!cols"] = [8, 11, 11, 12, 24, 10, 5, 10, 12, 14, 10, 15, 11, 24].map((wch) => ({ wch }));

    const stockRows = [
      ["Stock", day.toLocaleDateString()],
      [],
      ["Item", "Category", "Units left", ...locs.map((l) => `${l.name} price CHF`), `Value at ${locs[0].name} CHF`],
      ...data.products.map((p) => [
        p.name,
        p.type,
        stockOf(p.id),
        ...locs.map((l) => priceOf(p, l.id)),
        round2(stockOf(p.id) * priceOf(p, locs[0].id)),
      ]),
      [],
      [
        "TOTAL",
        "",
        data.products.reduce((s, p) => s + stockOf(p.id), 0),
        ...locs.map(() => ""),
        round2(data.products.reduce((s, p) => s + stockOf(p.id) * priceOf(p, locs[0].id), 0)),
      ],
    ];
    const stockSheet = XLSX.utils.aoa_to_sheet(stockRows);
    stockSheet["!cols"] = [{ wch: 24 }, { wch: 10 }, { wch: 11 }, ...locs.map(() => ({ wch: 16 })), { wch: 18 }];

    const settleRows = [
      ["Settlement with the common account", day.toLocaleDateString()],
      [],
      ["Person", "Holds money as", "Collected all time CHF", "Paid to common CHF", "Still owes CHF"],
      ...settleInfo.rows.map((a) => [a.name, a.method, round2(a.collected), round2(a.paid), a.due]),
      [],
      ["", "", "", "TOTAL OUTSTANDING", round2(settleInfo.outstanding)],
      [],
      ["Transfers"],
      ["Date", "Person", "Amount CHF", "Note"],
      ...data.transfers.map((t) => [new Date(t.ts).toLocaleDateString(), t.name, round2(t.amount), t.note || ""]),
    ];
    const settleSheet = XLSX.utils.aoa_to_sheet(settleRows);
    settleSheet["!cols"] = [20, 16, 21, 19, 16].map((wch) => ({ wch }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, salesSheet, "Sales");
    XLSX.utils.book_append_sheet(wb, stockSheet, "Stock");
    XLSX.utils.book_append_sheet(wb, settleSheet, "Settlement");
    try {
      XLSX.writeFile(wb, `Z-report-${isoDate(day)}.xlsx`);
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
          timeStr(s.ts),
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
    try {
      await navigator.clipboard.writeText(lines.join("\n"));
      setToast({ msg: "Copied — paste into a sheet." });
    } catch {
      setToast({ msg: "Copy blocked by the browser." });
    }
  };

  if (!authReady) return <div className="hl"><div className="empty">Opening the till…</div></div>;
  if (!session) return <SignIn />;
  if (!ready) return <div className="hl"><div className="empty">Opening the till…</div></div>;

  if (pickedProduct) {
    const listPrice = effectivePrice(pickedProduct);
    const available = stockOf(pickedProduct.id) - inCartQty(pickedProduct.id);
    const maxQty = Math.max(1, available);
    const previewPrice = priceMode === "gift" ? 0 : Number(unitPrice) || 0;
    return (
      <div className="hl">
        <button className="ghost" onClick={() => setPickedProduct(null)}>
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
        <button className="ghost solid wide mt16" onClick={() => addToCart(true)}>
          Confirm · CHF {money(cartTotal + qty * previewPrice)}
        </button>
        <button className="ghost wide mt8" onClick={() => addToCart(false)}>
          Add another item first
        </button>
      </div>
    );
  }

  if (tab === "pay") {
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
        <div className="cap mt16 mb6">Which account is it going to?</div>
        {data.accounts.length === 0 && <div className="empty">No accounts yet. Add them in Admin first.</div>}
        {data.accounts.map((a) => (
          <button key={a.id} className="who" onClick={() => checkout(a)}>
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
    return (
      <div className="hl">
        <button className="ghost" onClick={() => setTab("sell")}>
          ← Back
        </button>
        <div style={{ marginTop: 18 }}>
          <h1 className="xl">Start market day</h1>
        </div>
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
              items: data.products.map((p) => ({ pid: p.id, checked: false, price: priceOf(p, d.locId) })),
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
        save(
          { markets: data.markets.map((m) => (m.id === activeMarket.id ? { ...m, items: ticked } : m)) },
          [{ table: "markets", type: "update", id: activeMarket.id, row: { items: ticked } }]
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
          {isEditing ? "Save today's table" : "Start market"}
        </button>
      </div>
    );
  }

  if (summaryMarket) {
    const rows = data.sales.filter((s) => s.marketId === summaryMarket.id);
    const total = rows.reduce((s, x) => s + x.qty * x.price, 0);
    const units = rows.reduce((s, x) => s + x.qty, 0);
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
          Back to till
        </button>
      </div>
    );
  }

  if (activeMarket) {
    const marketSales = data.sales.filter((s) => s.marketId === activeMarket.id);
    const marketUnits = marketSales.reduce((s, x) => s + x.qty, 0);
    const marketCash = marketSales.reduce((s, x) => s + x.qty * x.price, 0);
    const marketProducts = activeMarket.items.map((item) => data.products.find((p) => p.id === item.pid)).filter(Boolean);
    return (
      <div className="hl">
        <div className="top">
          <div>
            <h1>{activeMarket.name}</h1>
            <div className="cap">Market day</div>
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

        {marketProducts.length === 0 && (
          <div className="empty">No products picked for today — use "Edit today's table" below.</div>
        )}
        <div className="tiles">
          {marketProducts.map((p) => {
            const available = stockOf(p.id) - inCartQty(p.id);
            return (
              <button key={p.id} className="tile" disabled={available <= 0} onClick={() => pickProduct(p)}>
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

        {cart.length > 0 && (
          <div className="bag">
            <div className="cap">This sale</div>
            {cart.map((l, i) => (
              <div className="line" key={i}>
                <span>
                  {l.qty}× {l.name}
                  {l.mode === "gift" ? " (gift)" : l.price < l.list ? " (reduced)" : ""}
                </span>
                <span>
                  <span className="num">{money(l.qty * l.price)}</span>
                  <button className="x" onClick={() => setCart(cart.filter((_, j) => j !== i))} aria-label="Remove line">
                    ×
                  </button>
                </span>
              </div>
            ))}
            <div className="tot">
              <span className="cap">Total</span>
              <span className="big">
                <small>CHF</small>
                {money(cartTotal)}
              </span>
            </div>
            <button className="ghost solid wide" onClick={() => setTab("pay")}>
              Note &amp; account →
            </button>
          </div>
        )}

        <hr className="rule" />
        <div className="cap">Today's sales</div>
        {marketSales.length === 0 && <div className="empty">Nothing sold yet.</div>}
        {marketSales.map((s) => (
          <div className="row" key={s.id}>
            <div className="grow">
              <div className="t">
                {s.qty}× {s.name}
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

        <hr className="rule" />
        <button
          className="ghost wide"
          onClick={() => {
            setMarketDraft({
              name: activeMarket.name,
              locId: activeMarket.locId,
              items: data.products.map((p) => {
                const found = activeMarket.items.find((i) => i.pid === p.id);
                return { pid: p.id, checked: !!found, price: found ? found.price : priceOf(p, activeMarket.locId) };
              }),
            });
            setTab("marketSetup2");
          }}
        >
          Edit today's table
        </button>
        <button className="ghost wide mt8" onClick={endMarket}>
          End market day
        </button>
      </div>
    );
  }

  return (
    <div className="hl">
      <div className="top">
        <div>
          <h1>BeeZness</h1>
          <div className="cap">{new Date().toLocaleDateString([], { weekday: "long", day: "numeric", month: "long" })}</div>
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
              <div className="tiles">
                {byType[type].map((p) => {
                  const available = stockOf(p.id) - inCartQty(p.id);
                  return (
                    <button key={p.id} className="tile" disabled={available <= 0} onClick={() => pickProduct(p)}>
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
            </div>
          ))}

          {cart.length > 0 && (
            <div className="bag">
              <div className="cap">This sale</div>
              {cart.map((l, i) => (
                <div className="line" key={i}>
                  <span>
                    {l.qty}× {l.name}
                    {l.mode === "gift" ? " (gift)" : l.price < l.list ? " (reduced)" : ""}
                  </span>
                  <span>
                    <span className="num">{money(l.qty * l.price)}</span>
                    <button className="x" onClick={() => setCart(cart.filter((_, j) => j !== i))} aria-label="Remove line">
                      ×
                    </button>
                  </span>
                </div>
              ))}
              <div className="tot">
                <span className="cap">Total</span>
                <span className="big">
                  <small>CHF</small>
                  {money(cartTotal)}
                </span>
              </div>
              <button className="ghost solid wide" onClick={() => setTab("pay")}>
                Note &amp; account →
              </button>
            </div>
          )}

          <hr className="rule" />
          <div className="cap">Today at {locationName(activeLoc)}</div>
          {today.list.length === 0 && <div className="empty">Nothing sold here yet today.</div>}
          {today.list.map((s) => (
            <div className="row" key={s.id}>
              <div className="grow">
                <div className="t">
                  {s.qty}× {s.name}
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
        </>
      )}

      {tab === "report" && (
        <>
          <div className="two">
            <div>
              <div className="cap mb6">Day</div>
              <select value={reportDay} onChange={(e) => setReportDay(e.target.value)}>
                {[...new Set([dayKey(Date.now()), ...saleDays])].map((d) => (
                  <option key={d} value={d}>
                    {new Date(d).toLocaleDateString([], { weekday: "short", day: "numeric", month: "long" })}
                  </option>
                ))}
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
          {Object.keys(report.byAcct).length === 0 && <div className="empty">Nothing on this day.</div>}
          {Object.entries(report.byAcct).map(([k, v]) => (
            <div className="row" key={k}>
              <div className="grow t">{k}</div>
              <div className="v">{money(v)}</div>
            </div>
          ))}

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
          <button className="ghost solid wide" onClick={downloadZReport}>
            Download Z report (Excel)
          </button>
          <button className="ghost wide mt8" onClick={copyCsv}>
            Copy as CSV instead
          </button>
          <div className="empty">
            Three sheets: the day's sales with notes, the stock you have left, and who still owes the common account.
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

      {tab === "admin" && (
        <>
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
                      const row = { id: uid(), pid: p.id, qty: n, ts: Date.now(), note: null };
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
                  onClick={() => save({ accounts: data.accounts.filter((x) => x.id !== a.id) }, [{ table: "accounts", type: "delete", id: a.id }])}
                  aria-label="Remove account"
                >
                  ×
                </button>
              </div>
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
          <div className="cap">Signed in as {session.user.email}</div>
          <button className="ghost mt8" onClick={() => supabase.auth.signOut()}>
            Sign out
          </button>
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
