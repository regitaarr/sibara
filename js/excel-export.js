// ============================================================
// excel-export.js
// Export Laporan BMN ke Excel dengan 3 Sheet terstruktur:
// 1. Neraca Barang (Rekapitulasi Total & Saldo)
// 2. Arus Masuk (Detail Barang Masuk & Foto Bukti Fisik)
// 3. Arus Keluar (Detail Barang Keluar & Foto Bukti Fisik)
// ============================================================

export async function exportToExcel(transaksiBulanData, bulanStr, allStok = [], allTransaksi = []) {
    if (!bulanStr) {
        throw new Error("Pilih bulan terlebih dahulu!");
    }

    const [yyyy, mm] = bulanStr.split("-").map(Number);
    const monthIndex = mm - 1;
    const year = yyyy;

    const monthNamesIndo = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ];
    const monthName = monthNamesIndo[monthIndex] || `Bulan_${mm}`;
    const monthNameUpper = monthName.toUpperCase();

    // ─── 1. Filter & Pisahkan Data Transaksi Bulan Terpilih ───────
    const sourceTxn = (allTransaksi && allTransaksi.length > 0) ? allTransaksi : transaksiBulanData;
    const monthTxn = sourceTxn.filter(t => {
        if (!t.tanggal_transaksi) return false;
        const d = t.tanggal_transaksi.toDate ? t.tanggal_transaksi.toDate() : new Date(t.tanggal_transaksi);
        if (isNaN(d.getTime())) return false;
        return d.getFullYear() === year && d.getMonth() === monthIndex;
    });

    // Urutkan transaksi secara kronologis (terlama ke terbaru)
    const sortedTxn = [...monthTxn].sort((a, b) => {
        const da = a.tanggal_transaksi?.toDate ? a.tanggal_transaksi.toDate() : new Date(a.tanggal_transaksi || 0);
        const db = b.tanggal_transaksi?.toDate ? b.tanggal_transaksi.toDate() : new Date(b.tanggal_transaksi || 0);
        return da - db;
    });

    // Kelompokkan data masuk dan data keluar
    const transaksiMasuk = sortedTxn.filter(t => t.jenis === "masuk");
    const transaksiKeluar = sortedTxn.filter(t => t.jenis === "keluar");

    // ─── 2. Rekapitulasi untuk Sheet 1 (Neraca Barang) ────────────
    const itemMap = new Map();

    // Inisialisasi dari master stok terkini
    allStok.forEach(item => {
        const barcode = item.id || item.kode_barcode || "-";
        itemMap.set(barcode, {
            barcode,
            nama: item.nama_barang || "-",
            satuan: item.satuan || "-",
            stokRealtime: Number(item.stok_terkini ?? 0),
            totalMasuk: 0,
            totalKeluar: 0
        });
    });

    // Akumulasi data mutasi transaksi bulan ini
    monthTxn.forEach(t => {
        const barcode = t.kode_barcode || t.id || "-";
        if (!itemMap.has(barcode)) {
            itemMap.set(barcode, {
                barcode,
                nama: t.nama_barang || "-",
                satuan: t.satuan || "-",
                stokRealtime: 0,
                totalMasuk: 0,
                totalKeluar: 0
            });
        }
        const rec = itemMap.get(barcode);
        const qty = Number(t.jumlah) || 0;
        if (t.jenis === "masuk") {
            rec.totalMasuk += qty;
        } else if (t.jenis === "keluar") {
            rec.totalKeluar += qty;
        }
    });

    const listRekap = Array.from(itemMap.values()).sort((a, b) =>
        a.nama.localeCompare(b.nama, "id")
    );

    // ─── 3. Inisialisasi Workbook ExcelJS ──────────────────────────
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "SIBARA - Sistem BMN";
    workbook.lastModifiedBy = "SIBARA";
    workbook.created = new Date();
    workbook.modified = new Date();

    // Style umum border tipis
    const thinBorder = {
        top: { style: "thin", color: { argb: "FFCBD5E1" } },
        left: { style: "thin", color: { argb: "FFCBD5E1" } },
        bottom: { style: "thin", color: { argb: "FFCBD5E1" } },
        right: { style: "thin", color: { argb: "FFCBD5E1" } }
    };

    const headerBorder = {
        top: { style: "thin", color: { argb: "FF94A3B8" } },
        left: { style: "thin", color: { argb: "FF94A3B8" } },
        bottom: { style: "medium", color: { argb: "FF475569" } },
        right: { style: "thin", color: { argb: "FF94A3B8" } }
    };

    const totalBorder = {
        top: { style: "thin", color: { argb: "FF475569" } },
        left: { style: "thin", color: { argb: "FFCBD5E1" } },
        bottom: { style: "double", color: { argb: "FF1E293B" } },
        right: { style: "thin", color: { argb: "FFCBD5E1" } }
    };

    // ============================================================
    // 📝 SHEET 1: Neraca Barang (Rekapitulasi Total)
    // ============================================================
    const ws1 = workbook.addWorksheet("Neraca Barang", {
        views: [{ showGridLines: true }],
        properties: { tabColor: { argb: "FF1E3A5F" } }
    });

    ws1.columns = [
        { key: "no", width: 6 },
        { key: "barcode", width: 18 },
        { key: "nama", width: 34 },
        { key: "satuan", width: 12 },
        { key: "masuk", width: 18 },
        { key: "keluar", width: 18 },
        { key: "saldoRentang", width: 22 },
        { key: "stokRealtime", width: 24 }
    ];

    // Judul & Header Info
    ws1.mergeCells("A1:H1");
    const title1 = ws1.getCell("A1");
    title1.value = `LAPORAN NERACA BARANG MILIK NEGARA (BMN)`;
    title1.font = { name: "Arial", size: 14, bold: true, color: { argb: "FF1E3A5F" } };
    title1.alignment = { vertical: "middle", horizontal: "center" };
    ws1.getRow(1).height = 26;

    ws1.mergeCells("A2:H2");
    const sub1 = ws1.getCell("A2");
    sub1.value = `Periode: ${monthNameUpper} ${year} | Dicetak: ${new Date().toLocaleDateString("id-ID", { day: "2-digit", month: "long", year: "numeric" })}`;
    sub1.font = { name: "Arial", size: 9.5, italic: true, color: { argb: "FF64748B" } };
    sub1.alignment = { vertical: "middle", horizontal: "center" };
    ws1.getRow(2).height = 18;

    ws1.getRow(3).height = 10; // Spacing row

    // Table Header Row 4
    const headers1 = [
        "No",
        "Barcode",
        "Nama Barang",
        "Satuan",
        "Jumlah Barang Masuk\n(Mutasi In)",
        "Jumlah Barang Keluar\n(Mutasi Out)",
        "Saldo Rentang Tanggal\n(Masuk - Keluar)",
        "Sisa Saldo Keseluruhan\n(Stok Real-time)"
    ];

    const hRow1 = ws1.getRow(4);
    hRow1.height = 34;
    headers1.forEach((text, idx) => {
        const cell = hRow1.getCell(idx + 1);
        cell.value = text;
        cell.font = { name: "Arial", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
        cell.border = headerBorder;

        // Color coding header
        if (idx === 4) {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0D9488" } }; // Teal Masuk
        } else if (idx === 5) {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDC2626" } }; // Red Keluar
        } else if (idx === 6) {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2563EB" } }; // Blue Saldo
        } else if (idx === 7) {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E3A5F" } }; // Dark Navy Stok
        } else {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF334155" } }; // Slate Header
        }
    });

    let rowIdx1 = 5;
    let sumMasuk = 0;
    let sumKeluar = 0;
    let sumSaldoRentang = 0;
    let sumStokRealtime = 0;

    listRekap.forEach((item, index) => {
        const saldoRentang = item.totalMasuk - item.totalKeluar;
        sumMasuk += item.totalMasuk;
        sumKeluar += item.totalKeluar;
        sumSaldoRentang += saldoRentang;
        sumStokRealtime += item.stokRealtime;

        const row = ws1.getRow(rowIdx1);
        row.height = 22;

        row.getCell(1).value = index + 1;
        row.getCell(2).value = item.barcode;
        row.getCell(3).value = item.nama;
        row.getCell(4).value = item.satuan;
        row.getCell(5).value = item.totalMasuk;
        row.getCell(6).value = item.totalKeluar;
        row.getCell(7).value = saldoRentang;
        row.getCell(8).value = item.stokRealtime;

        // Styling data cells
        for (let c = 1; c <= 8; c++) {
            const cell = row.getCell(c);
            cell.border = thinBorder;
            cell.font = { name: "Arial", size: 9 };
            if (c === 1 || c === 4) {
                cell.alignment = { vertical: "middle", horizontal: "center" };
            } else if (c === 2) {
                cell.alignment = { vertical: "middle", horizontal: "center" };
                cell.font = { name: "Consolas", size: 9 };
            } else if (c === 3) {
                cell.alignment = { vertical: "middle", horizontal: "left" };
                cell.font = { name: "Arial", size: 9, bold: true };
            } else if (c === 5 || c === 6 || c === 7 || c === 8) {
                cell.alignment = { vertical: "middle", horizontal: "right" };
                cell.numFmt = "#,##0";
            }
        }

        // Zebra background & highlight
        if (index % 2 === 1) {
            row.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
            row.getCell(2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
            row.getCell(3).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
            row.getCell(4).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
        }

        // Light tints for key numeric columns
        row.getCell(5).fill = { type: "pattern", pattern: "solid", fgColor: { argb: item.totalMasuk > 0 ? "FFF0FDF4" : "FFFFFFFF" } };
        row.getCell(6).fill = { type: "pattern", pattern: "solid", fgColor: { argb: item.totalKeluar > 0 ? "FFFEF2F2" : "FFFFFFFF" } };
        row.getCell(7).fill = { type: "pattern", pattern: "solid", fgColor: { argb: saldoRentang !== 0 ? "FFEFF6FF" : "FFFFFFFF" } };
        row.getCell(8).fill = { type: "pattern", pattern: "solid", fgColor: { argb: item.stokRealtime > 0 ? "FFF1F5F9" : "FFFFFFFF" } };

        rowIdx1++;
    });

    // Baris Total Sheet 1
    const totalRow1 = ws1.getRow(rowIdx1);
    totalRow1.height = 26;
    ws1.mergeCells(rowIdx1, 1, rowIdx1, 4);

    const totalLabelCell1 = totalRow1.getCell(1);
    totalLabelCell1.value = "TOTAL KESELURUHAN";
    totalLabelCell1.font = { name: "Arial", size: 9.5, bold: true, color: { argb: "FF0F172A" } };
    totalLabelCell1.alignment = { vertical: "middle", horizontal: "right" };

    totalRow1.getCell(5).value = sumMasuk;
    totalRow1.getCell(6).value = sumKeluar;
    totalRow1.getCell(7).value = sumSaldoRentang;
    totalRow1.getCell(8).value = sumStokRealtime;

    for (let c = 1; c <= 8; c++) {
        const cell = totalRow1.getCell(c);
        cell.border = totalBorder;
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
        if (c >= 5) {
            cell.font = { name: "Arial", size: 9.5, bold: true, color: { argb: "FF0F172A" } };
            cell.alignment = { vertical: "middle", horizontal: "right" };
            cell.numFmt = "#,##0";
        }
    }


    // ============================================================
    // 📥 SHEET 2: Arus Masuk (Detail Barang Masuk)
    // ============================================================
    const ws2 = workbook.addWorksheet("Arus Masuk", {
        views: [{ showGridLines: true }],
        properties: { tabColor: { argb: "FF10B981" } }
    });

    ws2.columns = [
        { key: "no", width: 6 },
        { key: "tanggal", width: 22 },
        { key: "barcode", width: 18 },
        { key: "nama", width: 32 },
        { key: "satuan", width: 12 },
        { key: "jumlah", width: 16 },
        { key: "inputBy", width: 26 },
        { key: "foto", width: 18 }
    ];

    // Header Judul Sheet 2
    ws2.mergeCells("A1:H1");
    const title2 = ws2.getCell("A1");
    title2.value = `BUKU PEMBANTU BARANG MASUK (ARUS MASUK)`;
    title2.font = { name: "Arial", size: 14, bold: true, color: { argb: "FF065F46" } };
    title2.alignment = { vertical: "middle", horizontal: "center" };
    ws2.getRow(1).height = 26;

    ws2.mergeCells("A2:H2");
    const sub2 = ws2.getCell("A2");
    sub2.value = `Periode: ${monthNameUpper} ${year} | Total Transaksi Masuk: ${transaksiMasuk.length}`;
    sub2.font = { name: "Arial", size: 9.5, italic: true, color: { argb: "FF64748B" } };
    sub2.alignment = { vertical: "middle", horizontal: "center" };
    ws2.getRow(2).height = 18;

    ws2.getRow(3).height = 10;

    // Headers Sheet 2
    const headers2 = [
        "No",
        "Tanggal Masuk",
        "Barcode",
        "Nama Barang",
        "Satuan",
        "Jumlah Masuk",
        "Input By",
        "Foto Bukti Fisik"
    ];

    const hRow2 = ws2.getRow(4);
    hRow2.height = 30;
    headers2.forEach((text, idx) => {
        const cell = hRow2.getCell(idx + 1);
        cell.value = text;
        cell.font = { name: "Arial", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.border = headerBorder;
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF059669" } }; // Emerald Green
    });

    let rowIdx2 = 5;
    let sumTotalMasukTxn = 0;

    if (transaksiMasuk.length === 0) {
        const emptyRow = ws2.getRow(rowIdx2);
        emptyRow.height = 28;
        ws2.mergeCells(rowIdx2, 1, rowIdx2, 8);
        const cell = emptyRow.getCell(1);
        cell.value = "Tidak ada transaksi barang masuk pada periode ini.";
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.font = { name: "Arial", size: 9.5, italic: true, color: { argb: "FF94A3B8" } };
        cell.border = thinBorder;
        rowIdx2++;
    } else {
        for (let i = 0; i < transaksiMasuk.length; i++) {
            const t = transaksiMasuk[i];
            const qty = Number(t.jumlah) || 0;
            sumTotalMasukTxn += qty;

            const row = ws2.getRow(rowIdx2);
            const hasPhoto = Boolean(t.foto_url);
            row.height = hasPhoto ? 65 : 24;

            row.getCell(1).value = i + 1;
            row.getCell(2).value = formatTanggalWaktu(t.tanggal_transaksi, t.created_at);
            row.getCell(3).value = t.kode_barcode || t.id || "-";
            row.getCell(4).value = t.nama_barang || "-";
            row.getCell(5).value = t.satuan || "-";
            row.getCell(6).value = qty;
            row.getCell(7).value = t.input_by || "-";

            for (let c = 1; c <= 8; c++) {
                const cell = row.getCell(c);
                cell.border = thinBorder;
                cell.font = { name: "Arial", size: 9 };
                if (c === 1 || c === 5) {
                    cell.alignment = { vertical: "middle", horizontal: "center" };
                } else if (c === 2) {
                    cell.alignment = { vertical: "middle", horizontal: "center" };
                } else if (c === 3) {
                    cell.alignment = { vertical: "middle", horizontal: "center" };
                    cell.font = { name: "Consolas", size: 9 };
                } else if (c === 4) {
                    cell.alignment = { vertical: "middle", horizontal: "left" };
                    cell.font = { name: "Arial", size: 9, bold: true };
                } else if (c === 6) {
                    cell.alignment = { vertical: "middle", horizontal: "right" };
                    cell.numFmt = "#,##0";
                    cell.font = { name: "Arial", size: 9, bold: true, color: { argb: "FF065F46" } };
                } else if (c === 7) {
                    cell.alignment = { vertical: "middle", horizontal: "left" };
                    cell.font = { name: "Arial", size: 8.5, color: { argb: "FF475569" } };
                } else if (c === 8) {
                    cell.alignment = { vertical: "middle", horizontal: "center" };
                }
            }

            // Tempel gambar bukti fisik jika tersedia
            if (hasPhoto) {
                await embedPhotoInCell(workbook, ws2, t.foto_url, 8, rowIdx2);
            } else {
                row.getCell(8).value = "—";
                row.getCell(8).font = { name: "Arial", size: 9, color: { argb: "FF94A3B8" } };
            }

            rowIdx2++;
        }

        // Total Row Sheet 2
        const totalRow2 = ws2.getRow(rowIdx2);
        totalRow2.height = 26;
        ws2.mergeCells(rowIdx2, 1, rowIdx2, 5);
        const lbl2 = totalRow2.getCell(1);
        lbl2.value = "TOTAL BARANG MASUK";
        lbl2.font = { name: "Arial", size: 9.5, bold: true, color: { argb: "FF065F46" } };
        lbl2.alignment = { vertical: "middle", horizontal: "right" };

        totalRow2.getCell(6).value = sumTotalMasukTxn;

        for (let c = 1; c <= 8; c++) {
            const cell = totalRow2.getCell(c);
            cell.border = totalBorder;
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD1FAE5" } };
            if (c === 6) {
                cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF065F46" } };
                cell.alignment = { vertical: "middle", horizontal: "right" };
                cell.numFmt = "#,##0";
            }
        }
    }


    // ============================================================
    // 📤 SHEET 3: Arus Keluar (Detail Barang Keluar)
    // ============================================================
    const ws3 = workbook.addWorksheet("Arus Keluar", {
        views: [{ showGridLines: true }],
        properties: { tabColor: { argb: "FFEF4444" } }
    });

    ws3.columns = [
        { key: "no", width: 6 },
        { key: "tanggal", width: 22 },
        { key: "barcode", width: 18 },
        { key: "nama", width: 32 },
        { key: "satuan", width: 12 },
        { key: "jumlah", width: 16 },
        { key: "inputBy", width: 26 },
        { key: "foto", width: 18 }
    ];

    // Header Judul Sheet 3
    ws3.mergeCells("A1:H1");
    const title3 = ws3.getCell("A1");
    title3.value = `BUKU PEMBANTU BARANG KELUAR (ARUS KELUAR)`;
    title3.font = { name: "Arial", size: 14, bold: true, color: { argb: "FF991B1B" } };
    title3.alignment = { vertical: "middle", horizontal: "center" };
    ws3.getRow(1).height = 26;

    ws3.mergeCells("A2:H2");
    const sub3 = ws3.getCell("A2");
    sub3.value = `Periode: ${monthNameUpper} ${year} | Total Transaksi Keluar: ${transaksiKeluar.length}`;
    sub3.font = { name: "Arial", size: 9.5, italic: true, color: { argb: "FF64748B" } };
    sub3.alignment = { vertical: "middle", horizontal: "center" };
    ws3.getRow(2).height = 18;

    ws3.getRow(3).height = 10;

    // Headers Sheet 3
    const headers3 = [
        "No",
        "Tanggal Keluar",
        "Barcode",
        "Nama Barang",
        "Satuan",
        "Jumlah Keluar",
        "Input By",
        "Foto Bukti Fisik"
    ];

    const hRow3 = ws3.getRow(4);
    hRow3.height = 30;
    headers3.forEach((text, idx) => {
        const cell = hRow3.getCell(idx + 1);
        cell.value = text;
        cell.font = { name: "Arial", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.border = headerBorder;
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDC2626" } }; // Ruby Red
    });

    let rowIdx3 = 5;
    let sumTotalKeluarTxn = 0;

    if (transaksiKeluar.length === 0) {
        const emptyRow = ws3.getRow(rowIdx3);
        emptyRow.height = 28;
        ws3.mergeCells(rowIdx3, 1, rowIdx3, 8);
        const cell = emptyRow.getCell(1);
        cell.value = "Tidak ada transaksi barang keluar pada periode ini.";
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.font = { name: "Arial", size: 9.5, italic: true, color: { argb: "FF94A3B8" } };
        cell.border = thinBorder;
        rowIdx3++;
    } else {
        for (let i = 0; i < transaksiKeluar.length; i++) {
            const t = transaksiKeluar[i];
            const qty = Number(t.jumlah) || 0;
            sumTotalKeluarTxn += qty;

            const row = ws3.getRow(rowIdx3);
            const hasPhoto = Boolean(t.foto_url);
            row.height = hasPhoto ? 65 : 24;

            row.getCell(1).value = i + 1;
            row.getCell(2).value = formatTanggalWaktu(t.tanggal_transaksi, t.created_at);
            row.getCell(3).value = t.kode_barcode || t.id || "-";
            row.getCell(4).value = t.nama_barang || "-";
            row.getCell(5).value = t.satuan || "-";
            row.getCell(6).value = qty;
            row.getCell(7).value = t.input_by || "-";

            for (let c = 1; c <= 8; c++) {
                const cell = row.getCell(c);
                cell.border = thinBorder;
                cell.font = { name: "Arial", size: 9 };
                if (c === 1 || c === 5) {
                    cell.alignment = { vertical: "middle", horizontal: "center" };
                } else if (c === 2) {
                    cell.alignment = { vertical: "middle", horizontal: "center" };
                } else if (c === 3) {
                    cell.alignment = { vertical: "middle", horizontal: "center" };
                    cell.font = { name: "Consolas", size: 9 };
                } else if (c === 4) {
                    cell.alignment = { vertical: "middle", horizontal: "left" };
                    cell.font = { name: "Arial", size: 9, bold: true };
                } else if (c === 6) {
                    cell.alignment = { vertical: "middle", horizontal: "right" };
                    cell.numFmt = "#,##0";
                    cell.font = { name: "Arial", size: 9, bold: true, color: { argb: "FF991B1B" } };
                } else if (c === 7) {
                    cell.alignment = { vertical: "middle", horizontal: "left" };
                    cell.font = { name: "Arial", size: 8.5, color: { argb: "FF475569" } };
                } else if (c === 8) {
                    cell.alignment = { vertical: "middle", horizontal: "center" };
                }
            }

            // Tempel gambar bukti fisik jika tersedia
            if (hasPhoto) {
                await embedPhotoInCell(workbook, ws3, t.foto_url, 8, rowIdx3);
            } else {
                row.getCell(8).value = "—";
                row.getCell(8).font = { name: "Arial", size: 9, color: { argb: "FF94A3B8" } };
            }

            rowIdx3++;
        }

        // Total Row Sheet 3
        const totalRow3 = ws3.getRow(rowIdx3);
        totalRow3.height = 26;
        ws3.mergeCells(rowIdx3, 1, rowIdx3, 5);
        const lbl3 = totalRow3.getCell(1);
        lbl3.value = "TOTAL BARANG KELUAR";
        lbl3.font = { name: "Arial", size: 9.5, bold: true, color: { argb: "FF991B1B" } };
        lbl3.alignment = { vertical: "middle", horizontal: "right" };

        totalRow3.getCell(6).value = sumTotalKeluarTxn;

        for (let c = 1; c <= 8; c++) {
            const cell = totalRow3.getCell(c);
            cell.border = totalBorder;
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } };
            if (c === 6) {
                cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF991B1B" } };
                cell.alignment = { vertical: "middle", horizontal: "right" };
                cell.numFmt = "#,##0";
            }
        }
    }


    // ─── 4. Eksekusi Download File Excel ──────────────────────────
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const fileName = `Laporan_BMN_${monthName}_${year}.xlsx`;

    if (typeof saveAs !== "undefined") {
        saveAs(blob, fileName);
    } else {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
}

// ─── Helper Functions ─────────────────────────────────────────

// Format tanggal & waktu: "12 Okt 2026, 09:00"
function formatTanggalWaktu(ts, fallbackTs) {
    const monthsShort = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
    const d = ts?.toDate ? ts.toDate() : (ts ? new Date(ts) : null);
    const fb = fallbackTs?.toDate ? fallbackTs.toDate() : (fallbackTs ? new Date(fallbackTs) : null);

    const baseDate = (d && !isNaN(d.getTime())) ? d : fb;
    if (!baseDate) return "-";

    const day = String(baseDate.getDate()).padStart(2, "0");
    const mon = monthsShort[baseDate.getMonth()];
    const yr = baseDate.getFullYear();

    let hours = baseDate.getHours();
    let minutes = baseDate.getMinutes();

    // Jika transaksi tanggal di-input format date-only (00:00), coba ambil jam dari timestamp pembuatan
    if (hours === 0 && minutes === 0 && fb && !isNaN(fb.getTime())) {
        hours = fb.getHours();
        minutes = fb.getMinutes();
    }

    const timeStr = `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
    return `${day} ${mon} ${yr}, ${timeStr}`;
}

// Menempelkan foto ke dalam sel tertentu secara proporsional
async function embedPhotoInCell(workbook, worksheet, photoUrl, colIndex, rowIndex) {
    try {
        let base64 = "";
        let ext = "jpeg";

        if (photoUrl.startsWith("data:image/")) {
            const parts = photoUrl.split(",");
            base64 = parts[1];
            if (photoUrl.includes("image/png")) ext = "png";
            else if (photoUrl.includes("image/webp")) ext = "png";
        } else {
            base64 = await fetchImageAsBase64(photoUrl);
            ext = detectExtension(photoUrl);
        }

        if (base64) {
            const imageId = workbook.addImage({ base64, extension: ext });
            worksheet.addImage(imageId, {
                tl: { col: colIndex - 1 + 0.05, row: rowIndex - 1 + 0.05 },
                br: { col: colIndex - 0.05, row: rowIndex - 0.05 },
                editAs: "oneCell"
            });
        }
    } catch (err) {
        console.warn("Gagal menempelkan gambar ke Excel:", err);
        worksheet.getRow(rowIndex).getCell(colIndex).value = "(Gagal Muat Foto)";
        worksheet.getRow(rowIndex).getCell(colIndex).font = { name: "Arial", size: 8, italic: true, color: { argb: "FF94A3B8" } };
    }
}

async function fetchImageAsBase64(url) {
    const response = await fetch(url, { mode: "cors" });
    if (!response.ok) throw new Error("Fetch gambar gagal: " + response.status);
    const blob = await response.blob();
    return await blobToBase64(blob);
}

function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result.split(",")[1]);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

function detectExtension(url) {
    const lower = url.toLowerCase();
    if (lower.includes(".png")) return "png";
    if (lower.includes(".webp")) return "png";
    return "jpeg";
}
