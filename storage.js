import { initializeApp } from "firebase/app";
import {
  getFirestore,
  doc,
  getDoc,
  setDoc,
  deleteDoc,
  collection,
  query,
  where,
  documentId,
  getDocs,
} from "firebase/firestore";
import { firebaseConfig } from "./firebaseConfig";

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// The app's storage calls look like:
//   window.storage.get(key, shared)
//   window.storage.set(key, value, shared)
//   window.storage.delete(key, shared)
//   window.storage.list(prefix, shared)
//
// "shared" data (tickets, employees, suggestions, presence) is data every signed-in
// employee needs to see -> stored in Firestore, in a collection called "storage".
// "personal" data (just the browser's remembered session:user) only needs to live in
// this one browser -> stored in localStorage, so we don't need real user auth for it.

const LOCAL_PREFIX = "flight_tickets_local:";

async function get(key, shared = false) {
  if (!shared) {
    const value = localStorage.getItem(LOCAL_PREFIX + key);
    return value === null ? null : { key, value, shared: false };
  }
  const snap = await getDoc(doc(db, "storage", key));
  if (!snap.exists()) return null;
  return { key, value: snap.data().value, shared: true };
}

async function set(key, value, shared = false) {
  if (!shared) {
    localStorage.setItem(LOCAL_PREFIX + key, value);
    return { key, value, shared: false };
  }
  await setDoc(doc(db, "storage", key), { value });
  return { key, value, shared: true };
}

async function del(key, shared = false) {
  if (!shared) {
    localStorage.removeItem(LOCAL_PREFIX + key);
    return { key, deleted: true, shared: false };
  }
  await deleteDoc(doc(db, "storage", key));
  return { key, deleted: true, shared: true };
}

async function list(prefix = "", shared = false) {
  if (!shared) {
    const keys = [];
    for (let i = 0; i < localStorage.length; i++) {
      const fullKey = localStorage.key(i);
      if (fullKey && fullKey.startsWith(LOCAL_PREFIX)) {
        const key = fullKey.slice(LOCAL_PREFIX.length);
        if (key.startsWith(prefix)) keys.push(key);
      }
    }
    return { keys, prefix, shared: false };
  }
  // Range query on the document ID to emulate a "starts with" prefix match.
  const col = collection(db, "storage");
  const q = query(
    col,
    where(documentId(), ">=", prefix),
    where(documentId(), "<", prefix + "\uf8ff")
  );
  const snap = await getDocs(q);
  const keys = [];
  snap.forEach((d) => keys.push(d.id));
  return { keys, prefix, shared: true };
}

window.storage = { get, set, delete: del, list };
