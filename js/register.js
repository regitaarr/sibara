// ============================================================
// register.js
// Handles user registration, form validation, and feedback
// ============================================================
import { registerUser } from "./auth.js";

document.addEventListener("DOMContentLoaded", () => {
    const form = document.getElementById("register-form");
    const alertBox = document.getElementById("alert-box");
    const btnRegister = document.getElementById("btn-register");
    const spinner = document.getElementById("spin-register");
    const btnText = btnRegister.querySelector(".btn-text");

    const namaInput = document.getElementById("nama");
    const emailInput = document.getElementById("email");
    const passwordInput = document.getElementById("password");
    const confirmPasswordInput = document.getElementById("confirm-password");

    // ─── Helpers ──────────────────────────────────────────────────
    function showAlert(msg, type = "danger") {
        alertBox.className = `alert alert-${type} py-2 mb-3`;
        alertBox.textContent = msg;
        alertBox.classList.remove("d-none");
    }

    function hideAlert() {
        alertBox.classList.add("d-none");
    }

    function setLoading(isLoading) {
        btnRegister.disabled = isLoading;
        spinner.classList.toggle("d-none", !isLoading);
        btnText.classList.toggle("d-none", isLoading);
    }

    // ─── Password visibility toggles ─────────────────────────────
    const togglePass = document.getElementById("toggle-pass");
    const eyeIcon = document.getElementById("eye-icon");
    togglePass.addEventListener("click", () => {
        const isPassword = passwordInput.type === "password";
        passwordInput.type = isPassword ? "text" : "password";
        eyeIcon.className = isPassword ? "bi bi-eye-slash text-muted" : "bi bi-eye text-muted";
    });

    const toggleConfirmPass = document.getElementById("toggle-confirm-pass");
    const eyeConfirmIcon = document.getElementById("eye-confirm-icon");
    toggleConfirmPass.addEventListener("click", () => {
        const isPassword = confirmPasswordInput.type === "password";
        confirmPasswordInput.type = isPassword ? "text" : "password";
        eyeConfirmIcon.className = isPassword ? "bi bi-eye-slash text-muted" : "bi bi-eye text-muted";
    });

    const roleSelect = document.getElementById("role");

    // ─── Form Submission ──────────────────────────────────────────
    form.addEventListener("submit", async (e) => {
        e.preventDefault();
        hideAlert();

        const nama = namaInput.value.trim();
        const email = emailInput.value.trim();
        const password = passwordInput.value;
        const confirmPassword = confirmPasswordInput.value;
        const role = roleSelect ? roleSelect.value : "user";

        // Validation
        if (!nama) {
            showAlert("Nama lengkap tidak boleh kosong.");
            namaInput.focus();
            return;
        }

        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!email || !emailRegex.test(email)) {
            showAlert("Masukkan alamat email yang valid.");
            emailInput.focus();
            return;
        }

        if (!password || password.length < 6) {
            showAlert("Password minimal harus 6 karakter.");
            passwordInput.focus();
            return;
        }

        if (password !== confirmPassword) {
            showAlert("Konfirmasi password tidak cocok.");
            confirmPasswordInput.focus();
            return;
        }

        if (!role) {
            showAlert("Silakan pilih role akun terlebih dahulu.");
            roleSelect.focus();
            return;
        }

        setLoading(true);

        try {
            await registerUser(nama, email, password, role);
            showAlert("Registrasi berhasil! Akun Anda sedang menunggu persetujuan dari Admin sebelum dapat digunakan untuk login. Mengalihkan ke halaman login...", "success");

            setTimeout(() => {
                window.location.href = "index.html";
            }, 3000);
        } catch (err) {
            console.error("Register Error:", err);
            let msg = "Terjadi kesalahan saat pendaftaran. Silakan coba lagi.";
            if (err.code === "auth/email-already-in-use") {
                msg = "Alamat email ini sudah terdaftar. Silakan gunakan email lain atau masuk.";
            } else if (err.code === "auth/invalid-email") {
                msg = "Format alamat email tidak valid.";
            } else if (err.code === "auth/weak-password") {
                msg = "Password terlalu lemah. Gunakan minimal 6 karakter.";
            } else if (err.code === "auth/network-request-failed") {
                msg = "Koneksi internet terputus. Silakan periksa jaringan Anda.";
            } else if (err.message) {
                msg = err.message;
            }
            showAlert(msg, "danger");
            setLoading(false);
        }
    });
});
