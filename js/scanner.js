// ============================================================
// scanner.js
// html5-qrcode scanner init, barcode handling, Firestore lookup.
// ============================================================
import { db } from "./firebase-config.js";
import { doc, getDoc } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

let html5QrCode = null;
let scanning = false;

const SCAN_CONFIG = {
    fps: 10,
    qrbox: { width: 250, height: 150 },
    aspectRatio: 1.0,
    supportedScanTypes: [
        window.Html5QrcodeScanType?.SCAN_TYPE_CAMERA ?? 0
    ]
};

// ─── Init / Start scanner ─────────────────────────────────────
export async function initScanner(elementId, onSuccess) {
    if (scanning) return;

    if (!html5QrCode) {
        html5QrCode = new Html5Qrcode(elementId);
    }

    try {
        await html5QrCode.start(
            { facingMode: "environment" },
            SCAN_CONFIG,
            (decodedText) => {
                if (scanning) return; // debounce
                scanning = true;
                onSuccess(decodedText);
            },
            () => { } // ignore scan errors (partial reads)
        );
    } catch (err) {
        console.error("Kamera gagal diinisialisasi:", err);
        throw err;
    }
}

// ─── Stop scanner ─────────────────────────────────────────────
export async function stopScanner() {
    if (html5QrCode) {
        try {
            await html5QrCode.stop();
        } catch (_) { }
    }
    scanning = false;
}

// ─── Restart scanner (after transaction) ─────────────────────
export async function restartScanner(elementId, onSuccess) {
    await stopScanner();
    // Brief pause so camera can be released properly
    await new Promise(r => setTimeout(r, 600));
    scanning = false;
    await initScanner(elementId, onSuccess);
}

// ─── Unlock scanner (allow next scan without restarting camera)
export function unlockScanner() {
    scanning = false;
}

// ─── Firestore: get barang by barcode ─────────────────────────
export async function getBarangByBarcode(barcode) {
    const ref = doc(db, "barang", barcode);
    const snap = await getDoc(ref);
    if (!snap.exists()) return null;
    return { id: snap.id, ...snap.data() };
}
