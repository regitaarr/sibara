// ============================================================
// transaction.js
// Image compression and Firestore atomic transaction.
// No Firebase Storage used. Image is saved as Base64 in Firestore.
// ============================================================
import { db } from "./firebase-config.js";
import {
    doc, setDoc, runTransaction, collection, serverTimestamp
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

// ─── Compress image (target < 200KB) ─────────────────────────
export async function compressImage(file) {
    const options = {
        maxSizeMB: 0.15, // Make it a bit smaller to be safe for Firestore limits
        maxWidthOrHeight: 800,
        useWebWorker: true,
        fileType: 'image/jpeg'
    };
    try {
        return await imageCompression(file, options);
    } catch (err) {
        console.warn("Kompresi gagal:", err);
        throw err;
    }
}

// ─── Get compressed photo as Base64 ──────────────────────────
export async function getBase64Photo(file) {
    if (!file) return "";
    const compressed = await compressImage(file);
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result); // Data URL includes 'data:image/jpeg;base64,...'
        reader.onerror = (e) => reject(e);
        reader.readAsDataURL(compressed);
    });
}

// ─── Save or update master barang (new item) ─────────────────
export async function saveMasterBarang(barcode, namaBarang, satuan) {
    const barangRef = doc(db, "barang", barcode);
    await setDoc(barangRef, {
        nama_barang: namaBarang,
        satuan: satuan,
        stok_terkini: 0,
        last_updated: serverTimestamp()
    }, { merge: true });
}

// ─── Atomic Firestore transaction ─────────────────────────────
// data: { kode_barcode, nama_barang, satuan, jenis, jumlah, tanggal, foto_url, input_by }
// Note: foto_url is now a base64 Data URL string
export async function saveTransaksi(data) {
    const { kode_barcode, nama_barang, satuan, jenis, jumlah, tanggal, foto_url, input_by } = data;

    const barangRef = doc(db, "barang", kode_barcode);
    const transaksiId = doc(collection(db, "transaksi")); // auto-id ref

    await runTransaction(db, async (txn) => {
        const barangSnap = await txn.get(barangRef);

        let stokSaat = 0;
        if (barangSnap.exists()) {
            stokSaat = barangSnap.data().stok_terkini ?? 0;
        }

        const stokBaru = jenis === "masuk"
            ? stokSaat + jumlah
            : stokSaat - jumlah;

        if (stokBaru < 0) {
            throw new Error(`Stok tidak mencukupi! Stok saat ini: ${stokSaat} ${satuan}`);
        }

        // Write new transaksi document
        txn.set(transaksiId, {
            kode_barcode,
            nama_barang,
            satuan,
            jenis,
            jumlah,
            tanggal_transaksi: tanggal
                ? new Date(tanggal)
                : new Date(),
            foto_url: foto_url || "",
            input_by,
            created_at: serverTimestamp()
        });

        // Update stok in barang collection
        if (barangSnap.exists()) {
            txn.update(barangRef, {
                stok_terkini: stokBaru,
                last_updated: serverTimestamp()
            });
        } else {
            txn.set(barangRef, {
                nama_barang,
                satuan,
                stok_terkini: stokBaru,
                last_updated: serverTimestamp()
            });
        }
    });
}
