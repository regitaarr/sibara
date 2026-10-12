// ============================================================
// excel-export.js
// Export Laporan BMN ke Excel dengan 3 Sheet terstruktur:
// 1. Neraca Barang (Matriks Tanggal Keluar & Masuk, Sabtu-Minggu Blok Hitam)
// 2. Arus Masuk (Detail Barang Masuk & Foto Bukti Fisik di Sel)
// 3. Arus Keluar (Detail Barang Keluar & Foto Bukti Fisik di Sel)
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
    const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();

    // ─── 1. Tentukan Hari Libur Akhir Pekan (Sabtu & Minggu) ───────
    const weekendDates = [];
    for (let d = 1; d <= daysInMonth; d++) {
        const date = new Date(year, monthIndex, d);
        if (date.getDay() === 0 || date.getDay() === 6) { // 0=Minggu, 6=Sabtu
            weekendDates.push(d);
        }
    }

    // ─── 2. Filter & Pisahkan Data Transaksi Bulan Terpilih ───────
    const sourceTxn = (allTransaksi && allTransaksi.length > 0) ? allTransaksi : transaksiBulanData;
    const monthTxn = sourceTxn.filter(t => {
        if (!t.tanggal_transaksi) return false;
        const d = t.tanggal_transaksi.toDate ? t.tanggal_transaksi.toDate() : new Date(t.tanggal_transaksi);
        if (isNaN(d.getTime())) return false;
        return d.getFullYear() === year && d.getMonth() === monthIndex;
    });

    // Urutkan transaksi secara kronologis
    const sortedTxn = [...monthTxn].sort((a, b) => {
        const da = a.tanggal_transaksi?.toDate ? a.tanggal_transaksi.toDate() : new Date(a.tanggal_transaksi || 0);
        const db = b.tanggal_transaksi?.toDate ? b.tanggal_transaksi.toDate() : new Date(b.tanggal_transaksi || 0);
        return da - db;
    });

    const transaksiMasuk = sortedTxn.filter(t => t.jenis === "masuk");
    const transaksiKeluar = sortedTxn.filter(t => t.jenis === "keluar");

    // ─── 3. Agregasi Matriks Tanggal per Barang untuk Sheet 1 ──────
    const itemStats = {};

    allStok.forEach(item => {
        const barcode = item.id || item.kode_barcode || "-";
        itemStats[barcode] = {
            barcode,
            nama: item.nama_barang || "-",
            satuan: item.satuan || "-",
            stokRealtime: Number(item.stok_terkini ?? 0),
            keluarDays: Array(daysInMonth).fill(0),
            masukDays: Array(daysInMonth).fill(0),
            totalKeluar: 0,
            totalMasuk: 0
        };
    });

    monthTxn.forEach(t => {
        const barcode = t.kode_barcode || t.id || "-";
        if (!itemStats[barcode]) {
            itemStats[barcode] = {
                barcode,
                nama: t.nama_barang || "-",
                satuan: t.satuan || "-",
                stokRealtime: 0,
                keluarDays: Array(daysInMonth).fill(0),
                masukDays: Array(daysInMonth).fill(0),
                totalKeluar: 0,
                totalMasuk: 0
            };
        }

        const d = t.tanggal_transaksi?.toDate ? t.tanggal_transaksi.toDate() : new Date(t.tanggal_transaksi);
        const day = d.getDate();
        const qty = Number(t.jumlah) || 0;

        // Pada hari Sabtu & Minggu tidak ada input/output
        if (weekendDates.includes(day)) {
            return;
        }

        if (day >= 1 && day <= daysInMonth) {
            if (t.jenis === "masuk") {
                itemStats[barcode].masukDays[day - 1] += qty;
                itemStats[barcode].totalMasuk += qty;
            } else if (t.jenis === "keluar") {
                itemStats[barcode].keluarDays[day - 1] += qty;
                itemStats[barcode].totalKeluar += qty;
            }
        }
    });

    const listRekap = Object.values(itemStats).sort((a, b) =>
        a.nama.localeCompare(b.nama, "id")
    );

    // Hitung total akumulasi per hari untuk baris "Jumlah Total ="
    const totalKeluarPerDay = Array(daysInMonth).fill(0);
    const totalMasukPerDay = Array(daysInMonth).fill(0);
    let grandTotalKeluar = 0;
    let grandTotalMasuk = 0;
    let grandTotalSaldoRentang = 0;
    let grandTotalStokRealtime = 0;

    listRekap.forEach(item => {
        for (let i = 0; i < daysInMonth; i++) {
            totalKeluarPerDay[i] += item.keluarDays[i];
            totalMasukPerDay[i] += item.masukDays[i];
        }
        grandTotalKeluar += item.totalKeluar;
        grandTotalMasuk += item.totalMasuk;
        grandTotalSaldoRentang += (item.totalMasuk - item.totalKeluar);
        grandTotalStokRealtime += item.stokRealtime;
    });

    // ─── 4. Inisialisasi Workbook ExcelJS ──────────────────────────
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "SIBARA - Sistem BMN";
    workbook.lastModifiedBy = "SIBARA";
    workbook.created = new Date();

    const thinBorder = {
        top: { style: "thin", color: { argb: "FFCBD5E1" } },
        left: { style: "thin", color: { argb: "FFCBD5E1" } },
        bottom: { style: "thin", color: { argb: "FFCBD5E1" } },
        right: { style: "thin", color: { argb: "FFCBD5E1" } }
    };

    const strongBorder = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } }
    };

    // ============================================================
    // 📝 SHEET 1: Neraca Barang (Matriks Tanggal Keluar & Masuk)
    // ============================================================
    const ws1 = workbook.addWorksheet("Neraca Barang", {
        views: [{ showGridLines: true }],
        properties: { tabColor: { argb: "FF1E3A5F" } }
    });

    // Kolom Matriks:
    // 1: No
    // 2: Barcode
    // 3: Nama Barang
    // 4: Satuan
    // 5 .. 4+days: Keluar per hari
    // 5+days: Jml Brg Keluar
    // 6+days .. 5+days*2: Masuk per hari
    // 6+days*2: Jml Brg Masuk
    // 7+days*2: Saldo Rentang Tanggal
    // 8+days*2: Sisa Saldo Keseluruhan (Stok Real-time)
    const keluarStartCol = 5;
    const keluarEndCol = 4 + daysInMonth;
    const jmlKeluarCol = 5 + daysInMonth;
    const masukStartCol = 6 + daysInMonth;
    const masukEndCol = 5 + daysInMonth * 2;
    const jmlMasukCol = 6 + daysInMonth * 2;
    const saldoRentangCol = 7 + daysInMonth * 2;
    const stokRealtimeCol = 8 + daysInMonth * 2;

    const cols1 = [
        { key: "no", width: 5 },
        { key: "barcode", width: 16 },
        { key: "nama", width: 28 },
        { key: "satuan", width: 8 }
    ];

    for (let d = 1; d <= daysInMonth; d++) {
        cols1.push({ key: `keluar_${d}`, width: 4.5 });
    }
    cols1.push({ key: "jml_keluar", width: 14 });

    for (let d = 1; d <= daysInMonth; d++) {
        cols1.push({ key: `masuk_${d}`, width: 4.5 });
    }
    cols1.push({ key: "jml_masuk", width: 14 });
    cols1.push({ key: "saldo_rentang", width: 18 });
    cols1.push({ key: "stok_realtime", width: 22 });

    ws1.columns = cols1;

    // Row 1 & 2 Headers
    const r1 = ws1.getRow(1);
    const r2 = ws1.getRow(2);
    r1.height = 25;
    r2.height = 24;

    // Merge identitas di Row 1 & 2
    ws1.mergeCells(1, 1, 2, 1); ws1.getCell("A1").value = "No.";
    ws1.mergeCells(1, 2, 2, 2); ws1.getCell("B1").value = "Barcode";
    ws1.mergeCells(1, 3, 2, 3); ws1.getCell("C1").value = "Nama Barang";
    ws1.mergeCells(1, 4, 2, 4); ws1.getCell("D1").value = "Satuan";

    // Header Keluar Group
    ws1.mergeCells(1, keluarStartCol, 1, keluarEndCol);
    r1.getCell(keluarStartCol).value = `Keluar ${monthNameUpper}`;

    // Header Jml Brg Keluar
    ws1.mergeCells(1, jmlKeluarCol, 2, jmlKeluarCol);
    r1.getCell(jmlKeluarCol).value = "Jml Brg Keluar";

    // Header Masuk Group
    ws1.mergeCells(1, masukStartCol, 1, masukEndCol);
    r1.getCell(masukStartCol).value = `Masuk ${monthNameUpper}`;

    // Header Jml Brg Masuk
    ws1.mergeCells(1, jmlMasukCol, 2, jmlMasukCol);
    r1.getCell(jmlMasukCol).value = "Jml Brg Masuk";

    // Header Saldo
    ws1.mergeCells(1, saldoRentangCol, 2, saldoRentangCol);
    r1.getCell(saldoRentangCol).value = "Saldo Rentang Tanggal";

    ws1.mergeCells(1, stokRealtimeCol, 2, stokRealtimeCol);
    r1.getCell(stokRealtimeCol).value = "Sisa Saldo Keseluruhan\n(Stok Real-time)";

    // Styling Headers Row 1
    const staticHeaderCols = [1, 2, 3, 4, jmlKeluarCol, jmlMasukCol, saldoRentangCol, stokRealtimeCol];
    staticHeaderCols.forEach(col => {
        const cell = r1.getCell(col);
        cell.font = { name: "Arial", size: 9.5, bold: true, color: { argb: "FF000000" } };
        cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        cell.border = strongBorder;
    });

    r1.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } };
    r1.getCell(2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } };
    r1.getCell(3).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } };
    r1.getCell(4).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } };

    // Warna grup Keluar & Masuk sesuai contoh gambar (Merah cerah)
    r1.getCell(keluarStartCol).font = { name: "Arial", size: 11, bold: true, color: { argb: "FF000000" } };
    r1.getCell(keluarStartCol).alignment = { horizontal: "center", vertical: "middle" };
    r1.getCell(keluarStartCol).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };
    r1.getCell(keluarStartCol).border = strongBorder;

    r1.getCell(jmlKeluarCol).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCA5A5" } }; // Light coral

    r1.getCell(masukStartCol).font = { name: "Arial", size: 11, bold: true, color: { argb: "FF000000" } };
    r1.getCell(masukStartCol).alignment = { horizontal: "center", vertical: "middle" };
    r1.getCell(masukStartCol).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };
    r1.getCell(masukStartCol).border = strongBorder;

    r1.getCell(jmlMasukCol).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCA5A5" } };
    r1.getCell(saldoRentangCol).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF7" } };
    r1.getCell(stokRealtimeCol).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF7" } };

    // Styling Tanggal di Row 2 (Keluar dan Masuk)
    for (let d = 1; d <= daysInMonth; d++) {
        const isWeekend = weekendDates.includes(d);

        // Kolom Keluar
        const cKeluar = r2.getCell(keluarStartCol + d - 1);
        cKeluar.value = isWeekend ? "" : d;
        cKeluar.alignment = { horizontal: "center", vertical: "middle" };
        cKeluar.border = strongBorder;

        if (isWeekend) {
            cKeluar.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF000000" } }; // BLOK HITAM
        } else {
            cKeluar.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } }; // Merah
            cKeluar.font = { name: "Arial", size: 9.5, bold: true, color: { argb: "FF000000" } };
        }

        // Kolom Masuk
        const cMasuk = r2.getCell(masukStartCol + d - 1);
        cMasuk.value = isWeekend ? "" : d;
        cMasuk.alignment = { horizontal: "center", vertical: "middle" };
        cMasuk.border = strongBorder;

        if (isWeekend) {
            cMasuk.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF000000" } }; // BLOK HITAM
        } else {
            cMasuk.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } }; // Merah
            cMasuk.font = { name: "Arial", size: 9.5, bold: true, color: { argb: "FF000000" } };
        }
    }

    // Border untuk cell yang di-merge di row 2
    [1, 2, 3, 4, jmlKeluarCol, jmlMasukCol, saldoRentangCol, stokRealtimeCol].forEach(c => {
        r2.getCell(c).border = strongBorder;
    });

    // ─── Row 3: Baris "Jumlah Total =" Sesuai Gambar Referensi ──────
    const r3 = ws1.getRow(3);
    r3.height = 24;

    ws1.mergeCells(3, 1, 3, 4);
    const lblTotal = r3.getCell(1);
    lblTotal.value = "Jumlah Total =";
    lblTotal.font = { name: "Arial", size: 9.5, bold: true };
    lblTotal.alignment = { horizontal: "right", vertical: "middle" };
    lblTotal.border = strongBorder;

    // Isi total Keluar per hari
    for (let d = 1; d <= daysInMonth; d++) {
        const colIdx = keluarStartCol + d - 1;
        const cell = r3.getCell(colIdx);
        cell.border = strongBorder;
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.font = { name: "Arial", size: 9, bold: true };

        if (weekendDates.includes(d)) {
            cell.value = "";
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF000000" } }; // BLOK HITAM
        } else {
            cell.value = totalKeluarPerDay[d - 1] > 0 ? totalKeluarPerDay[d - 1] : 0;
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } }; // Merah total
        }
    }

    // Total Jml Brg Keluar
    const cJmlKeluar = r3.getCell(jmlKeluarCol);
    cJmlKeluar.value = grandTotalKeluar;
    cJmlKeluar.font = { name: "Arial", size: 9.5, bold: true };
    cJmlKeluar.alignment = { horizontal: "right", vertical: "middle" };
    cJmlKeluar.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCA5A5" } };
    cJmlKeluar.border = strongBorder;
    cJmlKeluar.numFmt = "#,##0";

    // Isi total Masuk per hari
    for (let d = 1; d <= daysInMonth; d++) {
        const colIdx = masukStartCol + d - 1;
        const cell = r3.getCell(colIdx);
        cell.border = strongBorder;
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.font = { name: "Arial", size: 9, bold: true };

        if (weekendDates.includes(d)) {
            cell.value = "";
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF000000" } }; // BLOK HITAM
        } else {
            cell.value = totalMasukPerDay[d - 1] > 0 ? totalMasukPerDay[d - 1] : 0;
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };
        }
    }

    // Total Jml Brg Masuk
    const cJmlMasuk = r3.getCell(jmlMasukCol);
    cJmlMasuk.value = grandTotalMasuk;
    cJmlMasuk.font = { name: "Arial", size: 9.5, bold: true };
    cJmlMasuk.alignment = { horizontal: "right", vertical: "middle" };
    cJmlMasuk.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCA5A5" } };
    cJmlMasuk.border = strongBorder;
    cJmlMasuk.numFmt = "#,##0";

    // Total Saldo Rentang Tanggal
    const cSaldoRentang = r3.getCell(saldoRentangCol);
    cSaldoRentang.value = grandTotalSaldoRentang;
    cSaldoRentang.font = { name: "Arial", size: 9.5, bold: true };
    cSaldoRentang.alignment = { horizontal: "right", vertical: "middle" };
    cSaldoRentang.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF7" } };
    cSaldoRentang.border = strongBorder;
    cSaldoRentang.numFmt = "#,##0";

    // Total Sisa Stok Realtime
    const cStokRealtime = r3.getCell(stokRealtimeCol);
    cStokRealtime.value = grandTotalStokRealtime;
    cStokRealtime.font = { name: "Arial", size: 9.5, bold: true };
    cStokRealtime.alignment = { horizontal: "right", vertical: "middle" };
    cStokRealtime.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF7" } };
    cStokRealtime.border = strongBorder;
    cStokRealtime.numFmt = "#,##0";

    // ─── Row 4+: Data Barang ───────────────────────────────────────
    let rowIdx1 = 4;
    listRekap.forEach((item, index) => {
        const row = ws1.getRow(rowIdx1);
        row.height = 22;

        row.getCell(1).value = index + 1;
        row.getCell(2).value = item.barcode;
        row.getCell(3).value = item.nama;
        row.getCell(4).value = item.satuan;

        // Keluar per hari
        for (let d = 1; d <= daysInMonth; d++) {
            const colIdx = keluarStartCol + d - 1;
            const cell = row.getCell(colIdx);
            cell.border = thinBorder;
            cell.alignment = { horizontal: "center", vertical: "middle" };

            if (weekendDates.includes(d)) {
                cell.value = "";
                cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF000000" } }; // BLOK HITAM
            } else {
                const val = item.keluarDays[d - 1];
                cell.value = val > 0 ? val : "";
                cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
            }
        }

        // Jml Brg Keluar
        const cellKeluar = row.getCell(jmlKeluarCol);
        cellKeluar.value = item.totalKeluar > 0 ? item.totalKeluar : "";
        cellKeluar.border = thinBorder;
        cellKeluar.alignment = { horizontal: "right", vertical: "middle" };
        cellKeluar.font = { name: "Arial", size: 9, bold: item.totalKeluar > 0 };
        cellKeluar.fill = { type: "pattern", pattern: "solid", fgColor: { argb: item.totalKeluar > 0 ? "FFFEF2F2" : "FFFFFFFF" } };
        cellKeluar.numFmt = "#,##0";

        // Masuk per hari
        for (let d = 1; d <= daysInMonth; d++) {
            const colIdx = masukStartCol + d - 1;
            const cell = row.getCell(colIdx);
            cell.border = thinBorder;
            cell.alignment = { horizontal: "center", vertical: "middle" };

            if (weekendDates.includes(d)) {
                cell.value = "";
                cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF000000" } }; // BLOK HITAM
            } else {
                const val = item.masukDays[d - 1];
                cell.value = val > 0 ? val : "";
                cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
            }
        }

        // Jml Brg Masuk
        const cellMasuk = row.getCell(jmlMasukCol);
        cellMasuk.value = item.totalMasuk > 0 ? item.totalMasuk : "";
        cellMasuk.border = thinBorder;
        cellMasuk.alignment = { horizontal: "right", vertical: "middle" };
        cellMasuk.font = { name: "Arial", size: 9, bold: item.totalMasuk > 0 };
        cellMasuk.fill = { type: "pattern", pattern: "solid", fgColor: { argb: item.totalMasuk > 0 ? "FFF0FDF4" : "FFFFFFFF" } };
        cellMasuk.numFmt = "#,##0";

        // Saldo Rentang Tanggal
        const saldoRentang = item.totalMasuk - item.totalKeluar;
        const cellSaldoRentang = row.getCell(saldoRentangCol);
        cellSaldoRentang.value = saldoRentang;
        cellSaldoRentang.border = thinBorder;
        cellSaldoRentang.alignment = { horizontal: "right", vertical: "middle" };
        cellSaldoRentang.font = { name: "Arial", size: 9, bold: true };
        cellSaldoRentang.fill = { type: "pattern", pattern: "solid", fgColor: { argb: saldoRentang !== 0 ? "FFEFF6FF" : "FFFFFFFF" } };
        cellSaldoRentang.numFmt = "#,##0";

        // Sisa Saldo Real-time
        const cellStokRealtime = row.getCell(stokRealtimeCol);
        cellStokRealtime.value = item.stokRealtime;
        cellStokRealtime.border = thinBorder;
        cellStokRealtime.alignment = { horizontal: "right", vertical: "middle" };
        cellStokRealtime.font = { name: "Arial", size: 9, bold: true };
        cellStokRealtime.fill = { type: "pattern", pattern: "solid", fgColor: { argb: item.stokRealtime > 0 ? "FFF1F5F9" : "FFFFFFFF" } };
        cellStokRealtime.numFmt = "#,##0";

        // Basic info styling
        row.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
        row.getCell(1).border = thinBorder;
        row.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
        row.getCell(2).font = { name: "Consolas", size: 9 };
        row.getCell(2).border = thinBorder;
        row.getCell(3).alignment = { horizontal: "left", vertical: "middle" };
        row.getCell(3).font = { name: "Arial", size: 9, bold: true };
        row.getCell(3).border = thinBorder;
        row.getCell(4).alignment = { horizontal: "center", vertical: "middle" };
        row.getCell(4).border = thinBorder;

        rowIdx1++;
    });


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
        cell.border = strongBorder;
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF059669" } };
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

            if (hasPhoto) {
                await embedPhotoInCell(workbook, ws2, t.foto_url, 8, rowIdx2);
            } else {
                row.getCell(8).value = "—";
                row.getCell(8).font = { name: "Arial", size: 9, color: { argb: "FF94A3B8" } };
            }

            rowIdx2++;
        }

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
            cell.border = strongBorder;
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
        cell.border = strongBorder;
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDC2626" } };
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

            if (hasPhoto) {
                await embedPhotoInCell(workbook, ws3, t.foto_url, 8, rowIdx3);
            } else {
                row.getCell(8).value = "—";
                row.getCell(8).font = { name: "Arial", size: 9, color: { argb: "FF94A3B8" } };
            }

            rowIdx3++;
        }

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
            cell.border = strongBorder;
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } };
            if (c === 6) {
                cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF991B1B" } };
                cell.alignment = { vertical: "middle", horizontal: "right" };
                cell.numFmt = "#,##0";
            }
        }
    }


    // ─── 5. Eksekusi Download File Excel ──────────────────────────
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

    if (hours === 0 && minutes === 0 && fb && !isNaN(fb.getTime())) {
        hours = fb.getHours();
        minutes = fb.getMinutes();
    }

    const timeStr = `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
    return `${day} ${mon} ${yr}, ${timeStr}`;
}

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
