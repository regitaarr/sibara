// ============================================================
// users.js
// Admin-only: CRUD user records in Firestore `users` collection.
// Note: Creates Firestore records only. Firebase Auth account
// must be created separately (or via Firebase Console / Admin SDK).
// ============================================================
import { db } from "./firebase-config.js";
import {
    collection, doc, getDocs, setDoc, updateDoc, deleteDoc,
    query, orderBy, serverTimestamp
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

// ─── Load all users ───────────────────────────────────────────
export async function loadUsers() {
    const q = query(collection(db, "users"), orderBy("nama", "asc"));
    const snap = await getDocs(q);
    return snap.docs.map(d => ({ id: d.id, ...d.data() }));
}

// ─── Add / register user record ───────────────────────────────
export async function addUserRecord({ nama, email, role }) {
    const normalizedEmail = email.trim().toLowerCase();
    const ref = doc(db, "users", normalizedEmail);
    await setDoc(ref, {
        nama,
        role: role || "user",
        status: true,
        created_at: serverTimestamp()
    });
}

// ─── Update role ────────────────────────────────────────────────
export async function updateUserRole(email, role) {
    const ref = doc(db, "users", email.toLowerCase());
    await updateDoc(ref, { role });
}

// ─── Toggle status ────────────────────────────────────────────
export async function toggleUserStatus(email, currentStatus) {
    const ref = doc(db, "users", email.toLowerCase());
    await updateDoc(ref, { status: !currentStatus });
}

// ─── Delete user record (admin only) ─────────────────────────
export async function deleteUserRecord(email) {
    await deleteDoc(doc(db, "users", email.toLowerCase()));
}
