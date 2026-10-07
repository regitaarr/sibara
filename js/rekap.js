// ============================================================
// rekap.js
// Real-time stock listener and transaction history queries.
// ============================================================
import { db } from "./firebase-config.js";
import {
    collection, onSnapshot, query, orderBy, where, Timestamp,
    doc, deleteDoc, updateDoc, runTransaction, serverTimestamp
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

let unsubscribeStok = null;

// ─── Real-time stok listener ──────────────────────────────────
export function listenStok(onUpdate) {
    if (unsubscribeStok) unsubscribeStok();

    const q = query(collection(db, "barang"), orderBy("nama_barang", "asc"));
    unsubscribeStok = onSnapshot(q, (snapshot) => {
        const items = snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
        onUpdate(items);
    });

    return () => { if (unsubscribeStok) unsubscribeStok(); };
}

// ─── Load transaksi by month ──────────────────────────────────
// monthStr: "YYYY-MM" e.g. "2025-06"
export async function loadTransaksiByMonth(monthStr) {
    const [year, month] = monthStr.split("-").map(Number);
    const start = new Date(year, month - 1, 1, 0, 0, 0);
    const end = new Date(year, month, 0, 23, 59, 59);

    const q = query(
        collection(db, "transaksi"),
        where("tanggal_transaksi", ">=", Timestamp.fromDate(start)),
        where("tanggal_transaksi", "<=", Timestamp.fromDate(end)),
        orderBy("tanggal_transaksi", "asc")
    );

    return new Promise((resolve, reject) => {
        onSnapshot(q, (snapshot) => {
            const items = snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
            resolve(items);
        }, reject);
    });
}

// ─── Load all transaksi (no filter) ───────────────────────────
export function listenTransaksi(onUpdate) {
    const q = query(
        collection(db, "transaksi"),
        orderBy("tanggal_transaksi", "desc")
    );
    return onSnapshot(q, (snapshot) => {
        const items = snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
        onUpdate(items);
    });
}

// ─── Format Timestamp to readable date ───────────────────────
export function formatTanggal(ts) {
    if (!ts) return "-";
    const d = ts.toDate ? ts.toDate() : new Date(ts);
    return d.toLocaleDateString("id-ID", {
        day: "2-digit", month: "short", year: "numeric"
    });
}

// ─── Delete transaksi (admin only) ────────────────────────────
export async function deleteTransaksi(id) {
    await deleteDoc(doc(db, "transaksi", id));
}

// ─── Delete barang (admin only) ───────────────────────────────
export async function deleteBarang(barcode) {
    await deleteDoc(doc(db, "barang", barcode));
}

// ─── Update barang (user & admin) ─────────────────────────────
export async function updateBarang(barcode, newData) {
    const ref = doc(db, "barang", barcode);
    await updateDoc(ref, {
        nama_barang: newData.nama_barang,
        satuan: newData.satuan,
        last_updated: serverTimestamp()
    });
}

// ─── Update transaksi (user & admin) ──────────────────────────
export async function updateTransaksi(id, newData) {
    const tRef = doc(db, "transaksi", id);
    return await runTransaction(db, async (txn) => {
        const tSnap = await txn.get(tRef);
        if (!tSnap.exists()) throw new Error("Data transaksi tidak ditemukan!");
        const oldData = tSnap.data();

        // 1. Update stok in barang if jumlah changed
        const oldJumlah = oldData.jumlah || 0;
        const newJumlah = parseInt(newData.jumlah, 10);

        if (oldJumlah !== newJumlah) {
            const bRef = doc(db, "barang", oldData.kode_barcode);
            const bSnap = await txn.get(bRef);
            if (bSnap.exists()) {
                let stok = bSnap.data().stok_terkini || 0;

                // "Undo" old transaction
                if (oldData.jenis === "masuk") stok -= oldJumlah;
                else stok += oldJumlah;

                // Apply new transaction
                if (oldData.jenis === "masuk") stok += newJumlah;
                else stok -= newJumlah;

                if (stok < 0) {
                    throw new Error("Gagal: perubahan ini akan membuat stok menjadi negatif.");
                }

                txn.update(bRef, {
                    stok_terkini: stok,
                    last_updated: serverTimestamp()
                });
            }
        }

        // 2. Update transaction doc
        txn.update(tRef, {
            jumlah: newJumlah,
            tanggal_transaksi: new Date(newData.tanggal_transaksi)
            // Kita batasi hanya bisa edit jumlah & tanggal untuk menjaga konsistensi database
        });
    });
}
