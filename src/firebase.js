import { initializeApp } from "firebase/app";
import { getAuth, signInWithEmailAndPassword, signOut as fbSignOut, onAuthStateChanged, EmailAuthProvider, reauthenticateWithCredential, updatePassword, sendPasswordResetEmail } from "firebase/auth";
import {
  initializeFirestore,
  persistentLocalCache,
  persistentMultipleTabManager,
  collection,
  doc,
  onSnapshot,
  deleteDoc,
  setDoc,
  writeBatch,
} from "firebase/firestore";

// Konfigurace projektu vykazy-altepro
const firebaseConfig = {
  apiKey: "AIzaSyBU5sQ4Yf6iFjhuluY0wVptAdQv8cT_O4Q",
  authDomain: "vykazy-altepro.firebaseapp.com",
  projectId: "vykazy-altepro",
  storageBucket: "vykazy-altepro.firebasestorage.app",
  messagingSenderId: "411489128206",
  appId: "1:411489128206:web:2e5bb9d8f2047fae26515f",
};

const app = initializeApp(firebaseConfig);

// Firestore s offline cache — aplikace funguje i bez připojení,
// změny se synchronizují, jakmile je zařízení zpátky online.
export const db = initializeFirestore(app, {
  localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }),
});

// ---------- Přihlašování ----------
export const auth = getAuth(app);

export function subscribeAuth(callback) {
  return onAuthStateChanged(auth, (user) => callback(user || null));
}

export async function login(email, password) {
  await signInWithEmailAndPassword(auth, email, password);
}

export async function logout() {
  await fbSignOut(auth);
}

export async function changePassword(currentPassword, newPassword) {
  const user = auth.currentUser;
  if (!user || !user.email) throw new Error("Nikdo není přihlášen.");
  // Firebase vyžaduje čerstvé ověření před změnou hesla
  const cred = EmailAuthProvider.credential(user.email, currentPassword);
  await reauthenticateWithCredential(user, cred);
  await updatePassword(user, newPassword);
}

export async function resetPassword(email) {
  await sendPasswordResetEmail(auth, email);
}

// V testovací verzi se pracuje s oddělenými kolekcemi (test_people, test_entries, …),
// takže testování nikdy nesáhne na ostrá data. Přihlašovací účty jsou společné.
export const IS_TEST = import.meta.env.VITE_APP_ENV === "test";
const C = (name) => (IS_TEST ? "test_" + name : name);

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

// ---------- Odběry (živá synchronizace) ----------
export function subscribeCollection(name, callback, onError) {
  return onSnapshot(
    collection(db, C(name)),
    (snap) => callback(snap.docs.map((d) => ({ id: d.id, ...d.data() }))),
    (err) => onError && onError(err)
  );
}

export function subscribeTemplate(callback, onError) {
  return subscribeConfigDoc("template", callback, onError);
}

// Obecné sledování dokumentu v kolekci config (šablony apod.)
export function subscribeConfigDoc(id, callback, onError) {
  return onSnapshot(
    doc(db, C("config"), id),
    (snap) => callback(snap.exists() ? snap.data() : null),
    (err) => onError && onError(err)
  );
}

// ---------- Zápisy ----------
export async function addDocs(collName, items) {
  // Hromadný zápis — jedna dávka pro více dokumentů (např. výkaz pro celou směnu)
  const batch = writeBatch(db);
  items.forEach((item) => {
    const ref = doc(collection(db, C(collName)), uid());
    batch.set(ref, item);
  });
  await batch.commit();
}

export async function removeDoc(collName, id) {
  await deleteDoc(doc(db, C(collName), id));
}

export async function removeDocs(collName, ids) {
  // Hromadné mazání v jedné dávce
  const batch = writeBatch(db);
  ids.forEach((id) => batch.delete(doc(db, C(collName), id)));
  await batch.commit();
}

export async function saveTemplateDoc(meta) {
  await saveConfigDoc("template", meta);
}

export async function saveConfigDoc(id, data) {
  await setDoc(doc(db, C("config"), id), data);
}

export async function setUserDoc(email, data) {
  await setDoc(doc(db, C("users"), email), data);
}

// Vytvoří nebo přepíše dokument v libovolné kolekci (id = null -> nové)
export async function saveDocIn(collName, id, data) {
  const ref = id ? doc(db, C(collName), id) : doc(collection(db, C(collName)), uid());
  await setDoc(ref, data);
  return ref.id;
}
