// ============================================================
// auth.js
// Handles login, logout, password reset, and role-based routing.
// ============================================================
import { auth, db } from "./firebase-config.js";
import {
    signInWithEmailAndPassword,
    signOut,
    sendPasswordResetEmail,
    onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js";
import {
    doc,
    getDoc
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

// ─── Login ───────────────────────────────────────────────────
export async function loginUser(email, password) {
    const normalizedEmail = email.trim().toLowerCase();
    const credential = await signInWithEmailAndPassword(auth, normalizedEmail, password);
    return credential.user;
}

// ─── Logout ──────────────────────────────────────────────────
export async function logoutUser() {
    await signOut(auth);
    window.location.href = "index.html";
}

// ─── Password Reset ──────────────────────────────────────────
export async function sendResetEmail(email) {
    const normalizedEmail = email.trim().toLowerCase();
    await sendPasswordResetEmail(auth, normalizedEmail);
}

// ─── Get User Role from Firestore ────────────────────────────
export async function getUserRole(email) {
    const normalizedEmail = email.trim().toLowerCase();
    const userRef = doc(db, "users", normalizedEmail);
    const snap = await getDoc(userRef);
    if (!snap.exists()) throw new Error("Akun tidak ditemukan di sistem.");
    const data = snap.data();
    if (data.status === false) throw new Error("Akun Anda telah dinonaktifkan.");
    return data.role;
}

// ─── Auth Guard (call on protected pages) ────────────────────
// requiredRole: "admin" | "user" | null (any authenticated user)
export function requireAuth(requiredRole = null) {
    return new Promise((resolve, reject) => {
        const unsubscribe = onAuthStateChanged(auth, async (user) => {
            unsubscribe();
            if (!user) {
                window.location.href = "index.html";
                reject(new Error("Not authenticated"));
                return;
            }
            try {
                const role = await getUserRole(user.email);
                if (requiredRole === "admin" && role !== "admin") {
                    window.location.href = "dashboard.html";
                    reject(new Error("Not authorized"));
                    return;
                }
                resolve({ user, role });
            } catch (err) {
                await signOut(auth);
                window.location.href = "index.html";
                reject(err);
            }
        });
    });
}

// ─── Setup UI based on role ───────────────────────────────────
export function applyRoleUI(role) {
    const adminOnlyEls = document.querySelectorAll("[data-admin-only]");
    adminOnlyEls.forEach(el => {
        el.style.display = role === "admin" ? "" : "none";
    });
}
