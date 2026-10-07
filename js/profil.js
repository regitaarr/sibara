import { requireAuth, logoutUser, applyRoleUI } from "./auth.js";
import { db } from "./firebase-config.js";
import { doc, getDoc, updateDoc } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";
import { getAuth, updateProfile, updatePassword } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js";

// ─── Toast helper ─────────────────────────────────────────────
function showToast(msg, type = "success") {
    const el = document.createElement("div");
    el.className = `toast-bmn ${type}`;
    el.innerHTML = `<i class="bi bi-${type === "success" ? "check-circle" : type === "error" ? "x-circle" : "exclamation-triangle"} me-2"></i>${msg}`;
    document.getElementById("toast-container").appendChild(el);
    setTimeout(() => el.remove(), 5000);
}

let currentUser = null;
let currentRole = null;
const auth = getAuth();

document.addEventListener("DOMContentLoaded", async () => {
    // ─── Auth Guard ─────────────────────────────────────────────
    try {
        const authResult = await requireAuth("user");
        currentUser = authResult.user;
        currentRole = authResult.role;
        applyRoleUI(currentRole);
        loadProfileData();
    } catch (e) {
        // Redirected by requireAuth
    }

    // Toggle Password Visibility
    document.getElementById("toggle-pw").addEventListener("click", () => {
        const inp = document.getElementById("profil-password");
        const icon = document.getElementById("eye-icon");
        const isHidden = inp.type === "password";
        inp.type = isHidden ? "text" : "password";
        icon.className = isHidden ? "bi bi-eye-slash text-muted" : "bi bi-eye text-muted";
    });

    // Handle Form Submit
    document.getElementById("form-profil").addEventListener("submit", async (e) => {
        e.preventDefault();

        const alertEl = document.getElementById("profil-alert");
        alertEl.classList.add("d-none");

        const newNama = document.getElementById("profil-nama").value.trim();
        const newPassword = document.getElementById("profil-password").value;

        if (!newNama) {
            showAlert("Nama tidak boleh kosong.", "warning");
            return;
        }

        if (newPassword && newPassword.length < 6) {
            showAlert("Password baru minimal 6 karakter.", "warning");
            return;
        }

        const btn = document.getElementById("btn-save");
        const spin = document.getElementById("spin-save");
        btn.disabled = true;
        spin.classList.remove("d-none");
        btn.querySelector(".btn-text").classList.add("d-none");

        try {
            // Update Firestore Profile Name
            const userRef = doc(db, "users", currentUser.email);
            await updateDoc(userRef, { nama: newNama });

            // Update Auth Display Name
            await updateProfile(currentUser, { displayName: newNama });

            // Update Password if provided
            if (newPassword) {
                try {
                    await updatePassword(currentUser, newPassword);
                    document.getElementById("profil-password").value = "";
                } catch (pwErr) {
                    if (pwErr.code === "auth/requires-recent-login") {
                        throw new Error("Pembaruan password ditolak demi keamanan karena Anda sudah lama login. Silakan Logout dan login kembali untuk mengubah password.");
                    } else {
                        throw pwErr;
                    }
                }
            }

            showToast("Profil berhasil diperbarui!");
        } catch (err) {
            showAlert(err.message, "danger");
        } finally {
            btn.disabled = false;
            spin.classList.add("d-none");
            btn.querySelector(".btn-text").classList.remove("d-none");
        }
    });

    // Logout from header
    document.getElementById("btn-logout-header").addEventListener("click", () => {
        if (confirm("Keluar dari aplikasi?")) logoutUser();
    });

    // Logout from main content button
    document.getElementById("btn-logout").addEventListener("click", () => {
        if (confirm("Keluar dari aplikasi?")) logoutUser();
    });
});

async function loadProfileData() {
    if (!currentUser) return;

    document.getElementById("profil-email").value = currentUser.email;

    try {
        const userRef = doc(db, "users", currentUser.email);
        const snap = await getDoc(userRef);

        if (snap.exists() && snap.data().nama) {
            document.getElementById("profil-nama").value = snap.data().nama;
        } else if (currentUser.displayName) {
            document.getElementById("profil-nama").value = currentUser.displayName;
        }
    } catch (e) {
        console.error("Gagal memuat detail profil:", e);
    }
}

function showAlert(message, type) {
    const alertEl = document.getElementById("profil-alert");
    alertEl.textContent = message;
    alertEl.className = `alert alert-${type} py-2 mb-3`;
    alertEl.classList.remove("d-none");
}
