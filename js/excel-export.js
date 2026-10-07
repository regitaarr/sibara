// ============================================================
// excel-export.js (Advanced Daily Matrix Format)
// Modified to create day-by-day matrices, historical stock, and photos
// Tidy layout: dynamic row heights & vertical merges
// ============================================================

export async function exportToExcel(transaksiBulanData, bulanStr, allStok, allTransaksi) {
    if (!bulanStr) {
        throw new Error("Pilih bulan terlebih dahulu");
    }

    const [yyyy, mm] = bulanStr.split("-").map(Number);
    const monthIndex = mm - 1;
    const year = yyyy;

    // Get month name
    const monthNames = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"];
    const monthNameUpper = monthNames[monthIndex];
    const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Rekap Transaksi");

    workbook.creator = "Sistem BMN";
    workbook.created = new Date();

    // Calculate Weekends
    const weekendDates = [];
    for (let d = 1; d <= daysInMonth; d++) {
        const date = new Date(year, monthIndex, d);
        if (date.getDay() === 0 || date.getDay() === 6) {
            weekendDates.push(d);
        }
    }

    // Initialize Columns Array with proper widths
    const columns = [
        { header: "", key: "no", width: 5 }, // Headers will be manually set
        { header: "", key: "kd_barang", width: 14 },
        { header: "", key: "kd_brng", width: 10 },
        { header: "", key: "satuan", width: 8 },
        { header: "", key: "deskripsi", width: 28 }
    ];

    // Keluar Days
    for (let d = 1; d <= daysInMonth; d++) {
        columns.push({ header: d.toString(), key: `keluar_${d}`, width: 4.5 });
    }
    columns.push({ header: "Jml Brg Keluar", key: "jml_keluar", width: 14 });

    // Masuk Days
    for (let d = 1; d <= daysInMonth; d++) {
        columns.push({ header: d.toString(), key: `masuk_${d}`, width: 4.5 });
    }
    columns.push({ header: "Jml Brg Masuk", key: "jml_masuk", width: 14 });

    let monthShort = monthNameUpper.substring(0, 3);
    monthShort = monthShort[0] + monthShort.substring(1).toLowerCase(); // Sep
    columns.push({ header: `Saldo Barang per 1 ${monthShort}`, key: "saldo_awal", width: 14 });
    columns.push({ header: "Saldo Sisa Barang", key: "saldo_akhir", width: 14 });
    columns.push({ header: "DOKUMENTASI", key: "foto", width: 18 });

    worksheet.columns = columns;

    // Shift headers down to create Row 1 grouping
    worksheet.spliceRows(1, 0, []);

    // Setup Row 1 & 2 Headers and Merges
    const r1 = worksheet.getRow(1);
    const r2 = worksheet.getRow(2);
    r1.height = 25;
    r2.height = 25;

    // Merge A1:A2, B1:B2, C1:C2, D1:D2
    worksheet.mergeCells('A1:A2'); worksheet.getCell('A1').value = "No.";
    worksheet.mergeCells('B1:B2'); worksheet.getCell('B1').value = "Kd Barang";
    worksheet.mergeCells('C1:C2'); worksheet.getCell('C1').value = "Kd Brng";
    worksheet.mergeCells('D1:D2'); worksheet.getCell('D1').value = "Satuan";

    r1.getCell(5).value = "KETERANGAN";
    r2.getCell(5).value = "Deskripsi";

    const keluarStart = 6;
    const keluarEnd = 5 + daysInMonth;
    worksheet.mergeCells(1, keluarStart, 1, keluarEnd);
    r1.getCell(keluarStart).value = `Keluar ${monthNameUpper}`;

    const masukStart = 7 + daysInMonth;
    const masukEnd = 6 + daysInMonth * 2;
    worksheet.mergeCells(1, masukStart, 1, masukEnd);
    r1.getCell(masukStart).value = `Masuk ${monthNameUpper}`;

    const kuantitasCol = 7 + daysInMonth * 2; // AA
    r1.getCell(kuantitasCol).value = "KUANTITAS";

    const saldoAwalCol = 8 + daysInMonth * 2;
    const saldoAkhirCol = 9 + daysInMonth * 2;
    const fotoCol = 10 + daysInMonth * 2;

    worksheet.mergeCells(1, saldoAwalCol, 2, saldoAwalCol);
    worksheet.mergeCells(1, saldoAkhirCol, 2, saldoAkhirCol);
    worksheet.mergeCells(1, fotoCol, 2, fotoCol);

    // Styling Headers (Row 1 & 2)
    const headerCellsR1 = [1, 2, 3, 4, 5, keluarStart, masukStart, kuantitasCol, saldoAwalCol, saldoAkhirCol, fotoCol];
    headerCellsR1.forEach(col => {
        const cell = r1.getCell(col);
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });

    // Add background for R1 groups
    r1.getCell(keluarStart).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFAAAA' } };
    r1.getCell(masukStart).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFAAEEAA' } };
    r1.getCell(kuantitasCol).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } };

    // Apply borders and colors to all of R2
    r2.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        // Skip merged columns that take style from R1
        if ([1, 2, 3, 4, saldoAwalCol, saldoAkhirCol, fotoCol].includes(colNumber)) {
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            return;
        }

        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.font = { bold: true };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        let targetBg = 'FFF2F2F2';
        let dateDay = -1;

        if (colNumber >= 6 && colNumber <= 5 + daysInMonth) {
            dateDay = colNumber - 5;
        } else if (colNumber === 6 + daysInMonth) {
            targetBg = 'FFFCE4D6';
        } else if (colNumber >= 7 + daysInMonth && colNumber <= 6 + daysInMonth * 2) {
            dateDay = colNumber - (6 + daysInMonth);
        } else if (colNumber === 7 + daysInMonth * 2) {
            targetBg = 'FFC6E0B4';
        }

        if (dateDay !== -1 && weekendDates.includes(dateDay)) {
            targetBg = 'FF000000';
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        }

        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: targetBg } };
    });

    // Style the specific headers in R1/R2
    r1.getCell(saldoAwalCol).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
    r1.getCell(saldoAkhirCol).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
    r1.getCell(fotoCol).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };

    // Calculate Stats & Balances & Photos
    const itemStats = {};
    allStok.forEach((item, index) => {
        itemStats[item.id] = {
            itemInfo: item,
            indexNumber: index + 1,
            keluarDays: Array(daysInMonth).fill(0),
            masukDays: Array(daysInMonth).fill(0),
            jmlKeluar: 0,
            jmlMasuk: 0,
            saldoAwal: 0,
            saldoAkhir: 0,
            fotoUrl: ""
        };
    });

    const startOfMonth = new Date(year, monthIndex, 1, 0, 0, 0);

    // allTransaksi is sorted descending by date
    allTransaksi.forEach(t => {
        const stats = itemStats[t.kode_barcode];
        if (!stats) return;

        // Grab the most recent photo
        if (t.foto_url && !stats.fotoUrl) {
            stats.fotoUrl = t.foto_url;
        }

        if (!t.tanggal_transaksi) return;
        const tDate = t.tanggal_transaksi.toDate ? t.tanggal_transaksi.toDate() : new Date(t.tanggal_transaksi);

        if (tDate >= startOfMonth) {
            if (t.jenis === 'masuk') {
                stats.saldoAwal -= (t.jumlah || 0); // Reverse Masuk
            } else {
                stats.saldoAwal += (t.jumlah || 0); // Reverse Keluar
            }
        }
    });

    // Apply stok_terkini
    Object.values(itemStats).forEach(stats => {
        stats.saldoAwal += (stats.itemInfo.stok_terkini || 0);
    });

    // Aggregate monthly transactions
    transaksiBulanData.forEach(t => {
        if (!t.tanggal_transaksi) return;
        const tDate = t.tanggal_transaksi.toDate ? t.tanggal_transaksi.toDate() : new Date(t.tanggal_transaksi);
        const day = tDate.getDate();
        const idx = day - 1;

        const stats = itemStats[t.kode_barcode];
        if (stats) {
            if (t.jenis === 'masuk') {
                stats.masukDays[idx] += (t.jumlah || 0);
                stats.jmlMasuk += (t.jumlah || 0);
            } else {
                stats.keluarDays[idx] += (t.jumlah || 0);
                stats.jmlKeluar += (t.jumlah || 0);
            }
        }
    });

    let totalKeluarPerDay = Array(daysInMonth).fill(0);
    let totalMasukPerDay = Array(daysInMonth).fill(0);
    let totalJmlKeluar = 0;
    let totalJmlMasuk = 0;

    Object.values(itemStats).forEach(s => {
        s.saldoAkhir = s.saldoAwal + s.jmlMasuk - s.jmlKeluar;
        for (let i = 0; i < daysInMonth; i++) {
            totalKeluarPerDay[i] += s.keluarDays[i];
            totalMasukPerDay[i] += s.masukDays[i];
        }
        totalJmlKeluar += s.jmlKeluar;
        totalJmlMasuk += s.jmlMasuk;
    });

    // Initialize Row 3 (Totals)
    const row3Totals = { deskripsi: "Jumlah Total =" };
    for (let d = 1; d <= daysInMonth; d++) {
        row3Totals[`keluar_${d}`] = totalKeluarPerDay[d - 1] > 0 ? totalKeluarPerDay[d - 1] : "";
        row3Totals[`masuk_${d}`] = totalMasukPerDay[d - 1] > 0 ? totalMasukPerDay[d - 1] : "";
    }
    row3Totals.jml_keluar = totalJmlKeluar > 0 ? totalJmlKeluar : "";
    row3Totals.jml_masuk = totalJmlMasuk > 0 ? totalJmlMasuk : "";
    worksheet.addRow(row3Totals);

    const r3 = worksheet.getRow(3);
    r3.height = 25;

    // Merge A3 to E3 for neatness
    worksheet.mergeCells(3, 1, 3, 5);
    const totalLabelCell = r3.getCell(1);
    totalLabelCell.value = "Jumlah Total =";
    totalLabelCell.font = { bold: true };
    totalLabelCell.alignment = { horizontal: 'right', vertical: 'middle' };

    // Give all cells in R3 a border and background
    for (let c = 1; c <= fotoCol; c++) {
        const cell = r3.getCell(c);
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        if (c > 5) {
            cell.font = { bold: true };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            let bg = 'FFF2F2F2';
            let dateDay = -1;

            if (c >= 6 && c <= 5 + daysInMonth) {
                dateDay = c - 5;
            } else if (c === 6 + daysInMonth) {
                bg = 'FFFCE4D6';
            } else if (c >= 7 + daysInMonth && c <= 6 + daysInMonth * 2) {
                dateDay = c - (6 + daysInMonth);
            } else if (c === 7 + daysInMonth * 2) {
                bg = 'FFC6E0B4';
            } else if (c === 8 + daysInMonth * 2 || c === 9 + daysInMonth * 2) {
                bg = 'FFDDEBF7';
            }

            if (dateDay !== -1 && weekendDates.includes(dateDay)) {
                bg = 'FF000000';
            }
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        }
    }

    // Build Item Rows
    let rowIndex = 4;
    const ROW_HEIGHT_PHOTO = 60; // Just enough for a small thumbnail

    for (let itemIdx = 0; itemIdx < Object.values(itemStats).length; itemIdx++) {
        const stats = Object.values(itemStats)[itemIdx];
        const kdBrng = (itemIdx + 1).toString().padStart(6, '0');

        const rowData = {
            no: stats.indexNumber,
            kd_barang: stats.itemInfo.id || "",
            kd_brng: kdBrng,
            satuan: stats.itemInfo.satuan || "-",
            deskripsi: stats.itemInfo.nama_barang || "",
            jml_keluar: stats.jmlKeluar > 0 ? stats.jmlKeluar : "",
            jml_masuk: stats.jmlMasuk > 0 ? stats.jmlMasuk : "",
            saldo_awal: stats.saldoAwal > 0 ? stats.saldoAwal : "",
            saldo_akhir: stats.saldoAkhir > 0 ? stats.saldoAkhir : ""
        };

        for (let d = 1; d <= daysInMonth; d++) {
            rowData[`keluar_${d}`] = stats.keluarDays[d - 1] > 0 ? stats.keluarDays[d - 1] : "";
            rowData[`masuk_${d}`] = stats.masukDays[d - 1] > 0 ? stats.masukDays[d - 1] : "";
        }
        worksheet.addRow(rowData);

        const row = worksheet.getRow(rowIndex);
        // Dynamic row height for neatness: Only tall if it has a photo!
        row.height = stats.fotoUrl ? ROW_HEIGHT_PHOTO : 25;
        const hasStock = stats.saldoAkhir > 0;

        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            if (colNumber === 5) cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

            let bgColor = 'FFFFFFFF';
            let dateDay = -1;

            if ((colNumber <= 5 || (colNumber >= 8 + daysInMonth * 2 && colNumber < fotoCol)) && hasStock) {
                bgColor = 'FFC6E0B4';  // light green for items with stock
            }

            if (colNumber >= 6 && colNumber <= 5 + daysInMonth) {
                dateDay = colNumber - 5;
            } else if (colNumber === 6 + daysInMonth) {
                bgColor = 'FFFCE4D6';
            } else if (colNumber >= 7 + daysInMonth && colNumber <= 6 + daysInMonth * 2) {
                dateDay = colNumber - (6 + daysInMonth);
            } else if (colNumber === 7 + daysInMonth * 2) {
                bgColor = 'FFC6E0B4';
            }

            if (dateDay !== -1 && weekendDates.includes(dateDay)) {
                bgColor = 'FF000000';
            }

            if (bgColor !== 'FFFFFFFF') {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
            }
        });

        // Embed the most recent photo physically into the end column
        if (stats.fotoUrl) {
            try {
                let base64;
                let ext = "jpeg";

                if (stats.fotoUrl.startsWith("data:image/")) {
                    base64 = stats.fotoUrl.split(",")[1];
                    if (stats.fotoUrl.includes("image/png")) ext = "png";
                } else {
                    base64 = await fetchImageAsBase64(stats.fotoUrl);
                    ext = detectExtension(stats.fotoUrl);
                }

                if (base64) {
                    const imageId = workbook.addImage({ base64, extension: ext });
                    worksheet.addImage(imageId, {
                        tl: { col: fotoCol - 1, row: rowIndex - 1 },
                        br: { col: fotoCol, row: rowIndex },
                        editAs: "oneCell"
                    });
                }
            } catch (e) {
                row.getCell("foto").value = "(Gagal Muat Foto)";
            }
        }

        rowIndex++;
    }

    // Generate output
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const fileName = `Rekap_BMN_${monthNameUpper}_${year}.xlsx`;
    if (typeof saveAs !== 'undefined') {
        saveAs(blob, fileName);
    } else {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = fileName;
        link.click();
    }
}

// Helper: fetch URL -> Base64
async function fetchImageAsBase64(url) {
    const response = await fetch(url, { mode: "cors" });
    if (!response.ok) throw new Error("Fetch gagal: " + response.status);
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
