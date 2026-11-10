// Storage: Firebase (Auth + Firestore) when firebase-config.js is filled in, otherwise this device only (localStorage).
import { firebaseConfig } from "./firebase-config.js";

const FB = "https://www.gstatic.com/firebasejs/10.12.2";
export const isConfigured = () => firebaseConfig && firebaseConfig.apiKey && !firebaseConfig.apiKey.startsWith("PASTE");

const emptyState = () => ({ overrides: {}, balances: {}, paydays: {} });

// ---------------------------------------------------------------- local (testing / no Firebase yet)
function localStore(onChange) {
  const K = { model: "agm.model", state: "agm.state", entries: "agm.entries" };
  const rd = (k, d) => { try { return JSON.parse(localStorage.getItem(k)) ?? d; } catch { return d; } };
  const wr = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { alert("Could not save on this device: " + e.message); } };
  let data = { model: rd(K.model, null), state: rd(K.state, emptyState()), entries: rd(K.entries, []) };
  const emit = () => onChange({ ...data });
  setTimeout(emit, 0);
  return {
    mode: "local", user: { email: "this device only" },
    async setModel(m) { data.model = m; wr(K.model, m); emit(); },
    async setState(s) { data.state = s; wr(K.state, s); emit(); },
    async addEntry(e) { data.entries = [...data.entries, e]; wr(K.entries, data.entries); emit(); },
    async deleteEntry(id) { data.entries = data.entries.filter((e) => e.id !== id); wr(K.entries, data.entries); emit(); },
    async signOut() {},
  };
}

// ---------------------------------------------------------------- Firebase
async function firebaseStore(onChange, onAuth) {
  const { initializeApp } = await import(`${FB}/firebase-app.js`);
  const A = await import(`${FB}/firebase-auth.js`);
  const F = await import(`${FB}/firebase-firestore.js`);
  const app = initializeApp(firebaseConfig);
  const auth = A.getAuth(app);
  let db;
  try { db = F.initializeFirestore(app, { localCache: F.persistentLocalCache({ tabManager: F.persistentMultipleTabManager() }) }); }
  catch { db = F.getFirestore(app); }

  let unsubs = [];
  let data = { model: null, state: emptyState(), entries: [] };
  let uid = null;
  const api = {
    mode: "firebase", user: null,
    async signIn(email, pw) { await A.signInWithEmailAndPassword(auth, email, pw); },
    async signOut() { await A.signOut(auth); },
    async setModel(m) { await F.setDoc(F.doc(db, "users", uid, "app", "model"), m); },
    async setState(s) { await F.setDoc(F.doc(db, "users", uid, "app", "state"), s); },
    async addEntry(e) { await F.setDoc(F.doc(db, "users", uid, "entries", e.id), e); },
    async deleteEntry(id) { await F.deleteDoc(F.doc(db, "users", uid, "entries", id)); },
  };
  A.onAuthStateChanged(auth, (user) => {
    unsubs.forEach((u) => u()); unsubs = [];
    api.user = user; uid = user ? user.uid : null;
    onAuth(user);
    if (!user) return;
    data = { model: null, state: emptyState(), entries: [] };
    const emit = () => onChange({ ...data });
    unsubs.push(F.onSnapshot(F.doc(db, "users", uid, "app", "model"), (d) => { data.model = d.exists() ? d.data() : null; emit(); }));
    unsubs.push(F.onSnapshot(F.doc(db, "users", uid, "app", "state"), (d) => { data.state = d.exists() ? { ...emptyState(), ...d.data() } : emptyState(); emit(); }));
    unsubs.push(F.onSnapshot(F.collection(db, "users", uid, "entries"), (q) => { data.entries = q.docs.map((x) => x.data()); emit(); }));
  });
  return api;
}

export async function openStore(onChange, onAuth) {
  if (!isConfigured()) { const s = localStore(onChange); setTimeout(() => onAuth(s.user), 0); return s; }
  return firebaseStore(onChange, onAuth);
}
