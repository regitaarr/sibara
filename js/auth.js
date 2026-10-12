// ============================================================
// auth.js
// Handles login, logout, password reset, and role-based routing.
// ============================================================
import { auth, db } from "./firebase-config.js";
import {
    signInWithEmailAndPassword,
    createUserWithEmailAndPassword,
    updateProfile,
    signOut,
    sendPasswordResetEmail,
    onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js";
import {
    doc,
    getDoc,
    setDoc,
    serverTimestamp
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

// ─── Register ────────────────────────────────────────────────
export async function registerUser(nama, email, password, role = "user") {
    const normalizedEmail = email.trim().toLowerCase();
    const trimmedNama = nama.trim();
    const selectedRole = role || "user";

    // 1. Create account in Firebase Auth (or recover if created in previous attempt)
    let user;
    try {
        const credential = await createUserWithEmailAndPassword(auth, normalizedEmail, password);
        user = credential.user;
    } catch (authErr) {
        if (authErr.code === "auth/email-already-in-use") {
            try {
                const cred = await signInWithEmailAndPassword(auth, normalizedEmail, password);
                user = cred.user;
            } catch (_) {
                throw authErr;
            }
        } else {
            throw authErr;
        }
    }

    // 2. Set displayName in Firebase Auth
    try {
        await updateProfile(user, { displayName: trimmedNama });
    } catch (_) { }

    // 3. Save profile document in Firestore (status: false = pending admin approval)
    try {
        const userRef = doc(db, "users", normalizedEmail);
        await setDoc(userRef, {
            nama: trimmedNama,
            role: selectedRole,
            status: false,
            created_at: serverTimestamp()
        });
    } catch (fsErr) {
        console.warn("Firestore setDoc warning (Security rules may need update):", fsErr);
    }

    // 4. Sign out immediately so unapproved user cannot access app directly
    try {
        await signOut(auth);
    } catch (_) { }

    return user;
}

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
    try {
        const snap = await getDoc(userRef);
        if (snap.exists()) {
            const data = snap.data();
            if (data.status === false) {
                throw new Error("Akun Anda belum disetujui oleh Admin. Silakan hubungi Administrator.");
            }
            return data.role || "user";
        }
    } catch (e) {
        if (e.message.includes("belum disetujui") || e.message === "Akun Anda telah dinonaktifkan.") throw e;
        console.warn("Firestore getDoc warning:", e);
    }

    // If user document is missing or not yet approved in Firestore
    throw new Error("Akun Anda belum disetujui oleh Admin. Silakan hubungi Administrator.");
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
