
        import { requireAuth, logoutUser, applyRoleUI } from "./js/auth.js";
        import {
            listenStok, listenTransaksi, formatTanggal,
            deleteTransaksi, deleteBarang, updateBarang, updateTransaksi, renameBarang
        } from "./js/rekap.js";
        import { exportToExcel } from "./js/excel-export.js";

        // ─── Toast helper ─────────────────────────────────────────────
        function showToast(msg, type = "success") {
            const el = document.createElement("div");
            el.className = `toast-bmn ${type}`;
            el.innerHTML = `<i class="bi bi-${type === "success" ? "check-circle" : type === "error" ? "x-circle" : "exclamation-triangle"} me-2"></i>${msg}`;
            document.getElementById("toast-container").appendChild(el);
            setTimeout(() => el.remove(), 4000);
        }

        // ─── Auth ──────────────────────────────────────────────────────
        let currentRole = null;
        try {
            const { role } = await requireAuth();
            currentRole = role;
            applyRoleUI(role);

            document.getElementById("role-badge").textContent =
                role === "admin" ? "👑 Admin" : "👤 User";

            if (role === "admin") {
                document.getElementById("nav-users").classList.remove("d-none");
            }
        } catch (_) { }

        document.getElementById("btn-logout").addEventListener("click", () => {
            if (confirm("Keluar dari aplikasi?")) logoutUser();
        });

        // ─── Month filter default ──────────────────────────────────────
        const today = new Date();
        const ymStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}`;
        document.getElementById("filter-bulan").value = ymStr;

        // ─── Real-time stok ────────────────────────────────────────────
        let allStok = [];
        listenStok((items) => {
            allStok = items;
            document.getElementById("stat-total-item").textContent = items.length;
            document.getElementById("badge-stok-count").textContent = items.length;
            renderStok(items);
        });

        function renderStok(items) {
            const tbody = document.getElementById("tbody-stok");
            const isAdmin = currentRole === "admin";

            if (!items.length) {
                tbody.innerHTML = `<tr><td colspan="5">
        <div class="empty-state"><i class="bi bi-inbox"></i><p>Belum ada data barang.</p></div>
      </td></tr>`;
                return;
            }

            tbody.innerHTML = items.map((item, i) => `
      <tr>
        <td>${i + 1}</td>
        <td class="fw-600">${item.nama_barang || "—"}</td>
        <td>${item.satuan || "—"}</td>
        <td>
          <span class="pill ${item.stok_terkini === 0 ? "badge-keluar" : item.stok_terkini < 5 ? "" : "badge-masuk"}"
                style="${item.stok_terkini < 5 && item.stok_terkini > 0 ? "background:#FEF3C7;color:#92400E" : ""}">
            ${item.stok_terkini ?? 0}
          </span>
        </td>
        <td>
          <div class="d-flex gap-1 justify-content-center">
            <button class="btn btn-sm btn-outline-primary py-0 px-2"
                    onclick="openEditBarang('${item.id}', '${(item.nama_barang || "").replace(/'/g, "\\'")}', '${item.satuan}')"
                    title="Edit barang">
              <i class="bi bi-pencil"></i>
            </button>
            ${isAdmin ? `
            <button class="btn btn-sm btn-outline-danger py-0 px-2"
                    onclick="handleDeleteBarang('${item.id}', '${(item.nama_barang || "").replace(/'/g, "\\'")}')"
                    title="Hapus barang">
              <i class="bi bi-trash3"></i>
            </button>` : ""}
          </div>
        </td>
      </tr>
    `).join("");
        }

        // ─── Real-time transaksi ───────────────────────────────────────
        let allTransaksi = [];
        let filteredTransaksi = [];

        listenTransaksi((items) => {
            allTransaksi = items;
            applyMonthFilter();
        });

        function applyMonthFilter() {
            const bulan = document.getElementById("filter-bulan").value;
            if (!bulan) {
                filteredTransaksi = allTransaksi;
            } else {
                const [y, m] = bulan.split("-").map(Number);
                filteredTransaksi = allTransaksi.filter(t => {
                    const d = t.tanggal_transaksi?.toDate
                        ? t.tanggal_transaksi.toDate()
                        : new Date(t.tanggal_transaksi);
                    return d.getFullYear() === y && d.getMonth() + 1 === m;
                });
            }
            document.getElementById("stat-total-txn").textContent = filteredTransaksi.length;
            renderTransaksi(filteredTransaksi);
        }

        document.getElementById("btn-filter").addEventListener("click", applyMonthFilter);

        function renderTransaksi(items) {
            const tbody = document.getElementById("tbody-transaksi");
            const isAdmin = currentRole === "admin";
            document.getElementById("txn-info").textContent = `Menampilkan ${items.length} transaksi`;

            if (!items.length) {
                tbody.innerHTML = `<tr><td colspan="9">
        <div class="empty-state"><i class="bi bi-inbox"></i><p>Tidak ada transaksi pada periode ini.</p></div>
      </td></tr>`;
                return;
            }

            tbody.innerHTML = items.map((item, i) => {
                // YYYY-MM-DD format for input date value if needed
                const d = item.tanggal_transaksi?.toDate ? item.tanggal_transaksi.toDate() : new Date(item.tanggal_transaksi);
                const isoDate = d.toISOString().split("T")[0];

                return `
      <tr>
        <td>${i + 1}</td>
        <td class="text-nowrap">${formatTanggal(item.tanggal_transaksi)}</td>
        <td class="text-muted fs-sm">${item.kode_barcode || "—"}</td>
        <td class="fw-600">${item.nama_barang || "—"}</td>
        <td>
          <span class="pill badge-jenis badge-${item.jenis}">
            ${item.jenis === "masuk" ? "↓ Masuk" : "↑ Keluar"}
          </span>
        </td>
        <td class="text-end fw-600">${item.jumlah ?? 0} ${item.satuan || ""}</td>
        <td class="text-center">
          ${item.foto_url
                        ? `<button class="btn btn-sm btn-light border py-0 px-2 text-primary" onclick="showPhoto('${item.foto_url}')" title="Lihat Foto">
                <i class="bi bi-image"></i> Lihat
               </button>`
                        : `<span class="text-muted fs-sm">—</span>`}
        </td>
        <td class="fs-sm text-muted">${item.input_by || "—"}</td>
        <td>
          <div class="d-flex gap-1 justify-content-center flex-wrap">
            <button class="btn btn-sm btn-outline-primary py-0 px-2"
                    onclick="openEditTxn('${item.id}', '${isoDate}', ${item.jumlah}, '${(item.nama_barang || "").replace(/'/g, "\\'")}', '${item.jenis}')"
                    title="Edit transaksi">
              <i class="bi bi-pencil"></i>
            </button>
            ${isAdmin ? `
            <button class="btn btn-sm btn-outline-danger py-0 px-2"
                    onclick="handleDeleteTransaksi('${item.id}', '${(item.nama_barang || "").replace(/'/g, "\\'")}', '${item.jenis}')"
                    title="Hapus transaksi">
              <i class="bi bi-trash3"></i>
            </button>` : ""}
          </div>
        </td>
      </tr>
    `}).join("");
        }

        // ─── Modals ────────────────────────────────────────────────────
        const photoModal = new bootstrap.Modal(document.getElementById("modal-photo"));
        window.showPhoto = function (url) {
            document.getElementById("viewer-img").src = url;
            photoModal.show();
        };

        const editBarangModal = new bootstrap.Modal(document.getElementById("modal-edit-barang"));
        window.openEditBarang = function (id, nama, satuan) {
            document.getElementById("edit-barang-id").value = id;
            document.getElementById("edit-barang-barcode").value = id;
            document.getElementById("edit-barang-nama").value = nama;
            document.getElementById("edit-barang-satuan").value = satuan;
            editBarangModal.show();
        };

        document.getElementById("btn-save-barang").addEventListener("click", async () => {
            const oldId = document.getElementById("edit-barang-id").value;
            const newId = document.getElementById("edit-barang-barcode").value.trim();
            const data = {
                nama_barang: document.getElementById("edit-barang-nama").value.trim(),
                satuan: document.getElementById("edit-barang-satuan").value
            };
            if (!data.nama_barang || !newId) return showToast("Nama barang dan barcode wajib diisi.", "warning");

            const btn = document.getElementById("btn-save-barang");
            btn.disabled = true;
            try {
                if (oldId !== newId) {
                    await renameBarang(oldId, newId, data);
                    showToast(`Barcode / ID berhasil diubah dari ${oldId} menjadi ${newId}!`);
                } else {
                    await updateBarang(oldId, data);
                    showToast("Data barang berhasil diperbarui!");
                }
                editBarangModal.hide();
            } catch (err) {
                showToast(err.message || "Gagal mengupdate barang", "error");
            } finally {
                btn.disabled = false;
            }
        });

        const editTxnModal = new bootstrap.Modal(document.getElementById("modal-edit-txn"));
        window.openEditTxn = function (id, tanggal, jumlah, nama, jenis) {
            document.getElementById("edit-txn-id").value = id;
            document.getElementById("edit-txn-tanggal").value = tanggal;
            document.getElementById("edit-txn-jumlah").value = jumlah;
            document.getElementById("edit-txn-nama").textContent = nama;

            const badge = document.getElementById("edit-txn-jenis-badge");
            badge.textContent = jenis.toUpperCase();
            badge.className = `badge bg-${jenis === 'masuk' ? 'success' : 'danger'}`;

            editTxnModal.show();
        };

        document.getElementById("btn-save-txn").addEventListener("click", async () => {
            const id = document.getElementById("edit-txn-id").value;
            const data = {
                tanggal_transaksi: document.getElementById("edit-txn-tanggal").value,
                jumlah: document.getElementById("edit-txn-jumlah").value
            };

            if (!data.jumlah || data.jumlah < 1) return showToast("Jumlah tidak valid.", "warning");

            const btn = document.getElementById("btn-save-txn");
            btn.disabled = true;
            try {
                await updateTransaksi(id, data);
                showToast("Transaksi berhasil diperbarui!");
                editTxnModal.hide();
            } catch (err) {
                showToast(err.message || "Gagal mengupdate transaksi.", "error");
            } finally {
                btn.disabled = false;
            }
        });


        // ─── Delete handlers ───────────────────────────────────────────
        window.handleDeleteBarang = async function (barcode, nama) {
            if (!confirm(`Hapus barang "${nama}" (${barcode}) beserta semua riwayatnya dari master stok?\n\nCatatan: ini tidak menghapus riwayat transaksi.`)) return;
            try {
                await deleteBarang(barcode);
                showToast(`Barang "${nama}" berhasil dihapus dari master stok.`);
            } catch (err) {
                showToast("Gagal menghapus barang: " + err.message, "error");
            }
        };

        window.handleDeleteTransaksi = async function (id, nama, jenis) {
            if (!confirm(`Hapus transaksi "${jenis}" untuk "${nama}"?\n\nPeringatan: stok TIDAK otomatis dikembalikan.`)) return;
            try {
                await deleteTransaksi(id);
                showToast("Transaksi berhasil dihapus.");
            } catch (err) {
                showToast("Gagal menghapus transaksi: " + err.message, "error");
            }
        };

        // ─── Export Excel (ALL roles) ─────────────────────────────────
        document.getElementById("btn-export").addEventListener("click", async () => {
            if (!filteredTransaksi.length) {
                showToast("Tidak ada data untuk diekspor.", "warning");
                return;
            }
            const bulan = document.getElementById("filter-bulan").value || ymStr;
            const btn = document.getElementById("btn-export");
            const spin = document.getElementById("spin-export");
            btn.disabled = true;
            spin.classList.remove("d-none");

            try {
                await exportToExcel(filteredTransaksi, bulan, allStok, allTransaksi);
                showToast("File Excel berhasil diunduh!");
            } catch (err) {
                showToast("Gagal export Excel: " + err.message, "error");
            } finally {
                btn.disabled = false;
                spin.classList.add("d-none");
            }
        });
    