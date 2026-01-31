// 1. FUNGSI GLOBAL LOGIN (Tetap sama)
function handleCredentialResponse(response) {
    const idToken = response.credential;

    Swal.fire({
        title: 'Sedang Verifikasi...',
        allowOutsideClick: false,
        didOpen: () => { Swal.showLoading(); }
    });

    fetch("https://script.google.com/macros/s/AKfycbzlVEuazq6Sfcr8X_g5qdy75AQ5-vONvZBTPzfZxLtMtx9Zgpppd-9T_NmbJudyEt-E3g/exec", {
        method: "POST",
        body: JSON.stringify({
            action: "login",
            id_token: idToken
        })
    })
    .then(res => res.json())
    .then(data => {
        if (data.status === "ok") {
            // Simpan data user (email, nama, status, picture) ke session
            sessionStorage.setItem("user", JSON.stringify(data.user));
            
            Swal.fire({
                icon: 'success',
                title: 'Login Berhasil!',
                text: `Selamat datang, ${data.user.nama}`,
                timer: 2000,
                showConfirmButton: false
            }).then(() => { 
                window.location.reload(); 
            });
        } else {
            // Menampilkan pesan error dari GAS (misal: "Email anda belum diregistrasi!")
            Swal.fire('Login Gagal', data.message, 'error');
        }
    })
    .catch(err => {
        console.error(err);
        Swal.fire('Error', 'Terjadi kesalahan koneksi ke server autentikasi.', 'error');
    });
}

// 2. DOM CONTENT LOADED
document.addEventListener('DOMContentLoaded', () => {
    // State Aplikasi
    let viewDate = new Date(); // Bulan yang sedang dilihat di kalender
    let selectedDate = new Date(); // Tanggal yang dipilih untuk tabel
    let isMyBookingFilterActive = false;
    let allBookedData = [];
    let currentViewMode = 'daily';

    // Definisi Elemen
    const calendarUI = document.getElementById('calendar-ui');
    const tableBody = document.getElementById('tableBody');
    const selectRuang = document.getElementById('pilihRuang');
    const modal = document.getElementById('bookingModal');
    const btnOpen = document.getElementById('openModal');
    const authBtn = document.getElementById('authBtn');
    const form = document.getElementById('bookingForm');
    const selectJamMulai = document.getElementById('jamMulai');
    const selectJamSelesai = document.getElementById('jamSelesai');
    const monthLabel = document.getElementById('currentMonthLabel');
    const dateTitle = document.getElementById('selectedDateTitle');
    const tableLoading = document.getElementById('tableLoading');
    const tableTitle = document.getElementById('tableTitleDynamic');

    const jamOperasional = ["07:00", "08:00", "09:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00", "17:00"];
    const daftarRuang = ["RS.1", "RS.2", "RS.3", "Lab. Hidro", "Lab GGGF", "Lab GIIG", "Lab Foto", "Lab Geokom", "III.6", "III.5", "III.4", "III.3", "III.2", "III.1", "1.1", "R. Pengurus", "R. Sidang SURTA", "R. Bersama"];
    let filterRuangAktif = [...daftarRuang];

    // Admin
    const ADMIN_LIST = ["dwi.sapto@ugm.ac.id", 
                        "hanifahdwi@ugm.ac.id", 
                        "calvin.wijaya@mail.ugm.ac.id", 
                        "cecep.pratama@ugm.ac.id"];

    // Fungsi Utama Inisialisasi
    function initApp() {
        // Isi dropdown hanya sekali
        if (selectRuang.options.length <= 1) {
            selectRuang.innerHTML = '<option value="" disabled selected>Pilih Ruangan</option>';
            daftarRuang.forEach(ruang => {
                selectRuang.innerHTML += `<option value="${ruang}">${ruang}</option>`;
            });
            jamOperasional.forEach(jam => {
                selectJamMulai.innerHTML += `<option value="${jam}">${jam}</option>`;
                selectJamSelesai.innerHTML += `<option value="${jam}">${jam}</option>`;
            });
        }

        const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
        tableTitle.innerText = `Jadwal Penggunaan Ruang DTGD - ${new Date().toLocaleDateString('id-ID', options)}`;
        
        renderCalendar();
        renderDailyTable();
        checkLoginStatus();
    }

    // Menggambar Kalender di Card Kiri
    function renderCalendar() {
        const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        monthLabel.innerText = `${monthNames[viewDate.getMonth()]} ${viewDate.getFullYear()}`;

        const year = viewDate.getFullYear();
        const month = viewDate.getMonth();
        const firstDay = new Date(year, month, 1).getDay();
        const daysInMonth = new Date(year, month + 1, 0).getDate();

        let html = `<div class="calendar-grid">`;
        ['M', 'S', 'S', 'R', 'K', 'J', 'S'].forEach(d => html += `<small><b>${d}</b></small>`);

        for (let i = 0; i < firstDay; i++) html += `<div></div>`;

        for (let d = 1; d <= daysInMonth; d++) {
            const isSelected = (d === selectedDate.getDate() && month === selectedDate.getMonth() && year === selectedDate.getFullYear());
            const isToday = (d === new Date().getDate() && month === new Date().getMonth() && year === new Date().getFullYear());
            
            // --- LOGIKA PENANDA TITIK ---
            const dStr = String(d).padStart(2, '0');
            const mStr = String(month + 1).padStart(2, '0');
            const dateKey = `${dStr}/${mStr}/${year}`;
            
            // Cek apakah ada pesanan di tanggal ini
            const hasBooking = allBookedData.some(item => {
                const dateObj = new Date(item.tanggal);
                const itemD = String(dateObj.getDate()).padStart(2, '0');
                const itemM = String(dateObj.getMonth() + 1).padStart(2, '0');
                const itemY = dateObj.getFullYear();
                const itemFormatted = `${itemD}/${itemM}/${itemY}`;
                return itemFormatted === dateKey || item.tanggal === dateKey;
            });

            html += `
            <div class="cal-day ${isSelected ? 'cal-selected' : ''} ${isToday ? 'cal-today' : ''}" onclick="selectDate(${d})">
                ${d}
                ${hasBooking ? '<span class="cal-dot"></span>' : ''}
            </div>`;
        }
        html += `</div>`;
        calendarUI.innerHTML = html;
    }

    // Klik Tanggal di Kalender (Global agar bisa dipanggil dari HTML onclick)
    window.selectDate = (day) => {
        selectedDate = new Date(viewDate.getFullYear(), viewDate.getMonth(), day);
        
        const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
        const formattedDate = selectedDate.toLocaleDateString('id-ID', options);

        const titleElem = document.getElementById('tableTitleDynamic');
        if (titleElem) {
            titleElem.innerText = `Jadwal Penggunaan Ruang DTGD - ${formattedDate}`;
        }
        
        renderCalendar();

        if (currentViewMode === 'weekly') {
            renderWeeklyTable();
        } else {
            renderDailyTable();
        }
    };

    // Menggambar Tabel Harian (Satu Hari Saja)
    function renderDailyTable() {
        showTableLoading();
        tableBody.innerHTML = "";
        const y = selectedDate.getFullYear();
        const m = selectedDate.getMonth();
        const d = selectedDate.getDate();

        // Update Header Tabel (Th) di HTML agar sesuai filterRuangAktif
        const headerRow = document.querySelector('#agendaTable thead tr');
        headerRow.innerHTML = `<th>Jam</th>` + filterRuangAktif.map(r => `<th>${r}</th>`).join('');

        jamOperasional.forEach(jam => {
            let row = document.createElement('tr');
            let cells = `<td style="font-weight:bold; background:#f1f5f9;">${jam}</td>`;
            
            filterRuangAktif.forEach(r => {
                const idxAsli = daftarRuang.indexOf(r); // Mencari index asli untuk ID sel
                const id = `cell-${selectedDate.getFullYear()}-${selectedDate.getMonth()}-${selectedDate.getDate()}-${jam.replace(':','')}-${idxAsli}`;
                cells += `<td id="${id}"></td>`;
            });

            row.innerHTML = cells;
            tableBody.appendChild(row);
        });

        loadAndRenderAgenda();
    }

    // Logika Form Submit
    form.onsubmit = async (e) => {
        e.preventDefault();

        // 1. Persiapan Data
        const user = JSON.parse(sessionStorage.getItem("user") || "{}");
        const tglValue = document.getElementById('tglKegiatan').value;
        const tglInput = new Date(tglValue);

        // Ambil detail untuk cek tabrakan
        const ruang = selectRuang.value;
        const ruangIdx = daftarRuang.indexOf(ruang);
        const mulai = selectJamMulai.value;
        const selesai = selectJamSelesai.value;
        const idxMulai = jamOperasional.indexOf(mulai);
        const idxSelesai = jamOperasional.indexOf(selesai);

        // 2. Validasi Jam Dasar (Tetap sama)
        if (idxSelesai <= idxMulai) return Swal.fire('Error', 'Jam tidak valid!', 'error');

        // --- 2b. LOGIKA CEK TABRAKAN (CONFLICT CHECK) ---
        const y = tglInput.getFullYear();
        const m = tglInput.getMonth();
        const d = tglInput.getDate();
        
        let isConflict = false;
        for (let i = idxMulai; i < idxSelesai; i++) {
            const checkId = `cell-${y}-${m}-${d}-${jamOperasional[i].replace(':','')}-${ruangIdx}`;
            const targetCell = document.getElementById(checkId);
            
            // Jika sel di layar sudah mengandung class 'cell-filled', berarti bentrok
            if (targetCell && targetCell.classList.contains('cell-filled')) {
                isConflict = true;
                break;
            }
        }

        if (isConflict) {
            return Swal.fire({
                icon: 'error',
                title: 'Jadwal Bentrok!',
                text: 'Anda tidak bisa memesan ruangan di jam tersebut karena sudah ada agenda lainnya.',
                confirmButtonColor: '#e53e3e'
            });
        }
        
        // 3. Konfirmasi & Loading (Baru dijalankan jika tidak bentrok)
        const payload = {
            "Order ID": "DTGD-" + Date.now(),
            "Hari": tglInput.toLocaleDateString('id-ID', { weekday: 'long' }),
            "Tanggal": tglInput.toLocaleDateString('id-ID', { day: '2-digit', month: '2-digit', year: 'numeric' }),
            "Acara": document.getElementById('namaAcara').value,
            "Jam": `${selectJamMulai.value} - ${selectJamSelesai.value}`,
            "PIC Kegiatan": document.getElementById('picKegiatan').value,
            "Email PIC": user.email || "",
            "ruang": selectRuang.value
        };

        // 4. Konfirmasi & Loading
        Swal.fire({
            title: 'Kirim Pesanan?',
            text: `Memesan ${payload.ruang} untuk ${payload.Acara}`,
            icon: 'question',
            showCancelButton: true,
            confirmButtonText: 'Ya, Kirim'
        }).then(async (res) => {
            if (!res.isConfirmed) return;

            Swal.fire({ title: 'Menyimpan...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });

            try {
                // --- PENDEKATAN FORMDATA ---
                const fd = new FormData();
                fd.append("data", JSON.stringify(payload));

                const response = await fetch("https://script.google.com/macros/s/AKfycbxQ8poOfZ9y4dZVEOQ0qxWxBJ9i8P0-HlHkhCEBYWpooP7kHobijeUcHc6uHCc6uyN5/exec", {
                    method: "POST",
                    body: fd // Kirim sebagai FormData
                });

                const result = await response.json();

                if (result.result === "success") {
                    // Visualisasi ke Tabel
                    renderVisualAgenda(tglInput, idxMulai, idxSelesai, payload.ruang, payload.Acara, payload["PIC Kegiatan"], payload["Email PIC"], payload["Order ID"]);
                    
                    modal.style.display = "none";
                    form.reset();
                    Swal.fire('Berhasil!', 'Data tersimpan di Cloud Spreadsheet.', 'success');
                } else {
                    throw new Error(result.message);
                }

            } catch (err) {
                console.error(err);
                Swal.fire('Gagal Simpan', 'Terjadi kesalahan koneksi atau CORS. Pastikan GAS sudah di-deploy sebagai "Anyone".', 'error');
            }
        });
    };

    // Fungsi pembantu untuk menggambar di tabel setelah sukses
    function renderVisualAgenda(dateObj, startIdx, endIdx, ruang, acara, pic, emailPIC = "", orderID = "") {
        const user = JSON.parse(sessionStorage.getItem("user") || "{}");
        const userEmail = user.email ? user.email.toLowerCase().trim() : "";
    
        const isOwner = userEmail === (emailPIC ? emailPIC.toLowerCase().trim() : "");
        const isAdmin = ADMIN_LIST.includes(userEmail);
        const hasAccess = isOwner || isAdmin;

        const dataAsli = allBookedData.find(item => item.orderId === orderID);
    
        const fullData = dataAsli || {
            tanggal: `${String(dateObj.getDate()).padStart(2,'0')}/${String(dateObj.getMonth()+1).padStart(2,'0')}/${dateObj.getFullYear()}`,
            ruang, 
            acara, 
            pic, 
            jam: `${jamOperasional[startIdx]} - ${jamOperasional[endIdx]}`, // Gunakan endIdx tanpa -1
            emailPIC, 
            orderId: orderID
        };

        const cellBaseId = `cell-${dateObj.getFullYear()}-${dateObj.getMonth()}-${dateObj.getDate()}`;

        const y = dateObj.getFullYear();
        const m = dateObj.getMonth();
        const d = dateObj.getDate();
        const ruangIdx = daftarRuang.indexOf(ruang);
        const randomColor = ['#3182ce', '#38a169', '#d69e2e', '#e53e3e', '#805ad5'][Math.floor(Math.random() * 5)];

        // Logika Warna
        let finalColor = ['#3182ce', '#38a169', '#d69e2e', '#e53e3e', '#805ad5'][Math.floor(Math.random() * 5)];
        let opacity = "1";

        if (isMyBookingFilterActive) {
            if (!hasAccess) {
                finalColor = "#cbd5e0"; // Abu-abu jika bukan milik saya
                opacity = "0.5";
            } else {
                finalColor = "#2b6cb0"; // Biru pekat milik saya
            }
        }

        for (let i = startIdx; i < endIdx; i++) {
            const cellId = `cell-${y}-${m}-${d}-${jamOperasional[i].replace(':','')}-${ruangIdx}`;
            const cell = document.getElementById(cellId);
            if (cell) {
                cell.classList.add('cell-filled');
                let pos = "agenda-mid";
                if (i === startIdx) pos = "agenda-start";
                if (i === endIdx - 1) pos = "agenda-end";
                
                const block = document.createElement('div');
                block.className = `agenda-block ${pos}`;
                
                block.style.backgroundColor = finalColor; 
                block.style.opacity = opacity;
                
                block.innerHTML = (i === startIdx) ? `<strong>${acara}</strong><br><small>${pic}</small>` : '';
                block.onclick = () => showDetailModal(fullData);
                
                cell.innerHTML = "";
                cell.appendChild(block);
            }
        }
    }

    // Navigasi Bulan
    document.getElementById('prevMonth').onclick = () => { viewDate.setMonth(viewDate.getMonth() - 1); renderCalendar(); };
    document.getElementById('nextMonth').onclick = () => { viewDate.setMonth(viewDate.getMonth() + 1); renderCalendar(); };

    // Login & Modal UI
    function checkLoginStatus() {
        const userJson = sessionStorage.getItem("user");
        const btnMyBooking = document.getElementById('btnMyBooking');

        if (userJson) {
            const user = JSON.parse(userJson);
            btnOpen.classList.remove('disabled');
            btnMyBooking.classList.remove('disabled');

            btnMyBooking.style.background = "";
            btnMyBooking.style.color = "";  

            authBtn.textContent = `Logout (${user.nama.split(' ')[0]})`;
            authBtn.onclick = () => {
                sessionStorage.removeItem("user");
                window.location.reload();
            };
        } else {
            btnOpen.classList.add('disabled');
            btnMyBooking.classList.add('disabled');

            btnMyBooking.style.background = ""; 
            btnMyBooking.style.color = "";

            authBtn.textContent = "Login";
            authBtn.onclick = () => {
                Swal.fire({
                    title: 'Login ke Pesan Ruang',
                    html: `<div style="display:flex;justify-content:center;margin-top:20px;"><div id="google-login-target"></div></div>`,
                    showConfirmButton: false,
                    showCloseButton: true,
                    didOpen: () => {
                        google.accounts.id.renderButton(document.getElementById("google-login-target"), { theme: "outline", size: "large", shape: "pill" });
                    }
                });
            };
        }
    }

    // Modal Control
    btnOpen.onclick = () => {
        if (btnOpen.classList.contains('disabled')) {
            Swal.fire('Akses Terbatas', 'Silakan Login terlebih dahulu.', 'warning');
        } else {
            modal.style.display = "block";
            const user = JSON.parse(sessionStorage.getItem("user"));
            if(user) document.getElementById('picKegiatan').value = user.nama || "";
        }
    };
    document.querySelector('.close').onclick = () => modal.style.display = "none";

    // Start App
    initApp();

    // Fungsi untuk memanggil data dari spreadsheet
    async function loadAndRenderAgenda() {
        const d = String(selectedDate.getDate()).padStart(2, '0');
        const m = String(selectedDate.getMonth() + 1).padStart(2, '0');
        const y = selectedDate.getFullYear();
        const targetDateStr = `${d}/${m}/${y}`;
        
        try {
            const response = await fetch("https://script.google.com/macros/s/AKfycbxQ8poOfZ9y4dZVEOQ0qxWxBJ9i8P0-HlHkhCEBYWpooP7kHobijeUcHc6uHCc6uyN5/exec?t=" + Date.now());
            allBookedData = await response.json()

            const dailyAgenda = allBookedData.filter(item => {
                let itemFormatted = "";

                if (item.tanggal.includes('/') && item.tanggal.length <= 10) {
                    // Kasus A: Data sudah berupa string "31/01/2026"
                    itemFormatted = item.tanggal;
                } else {
                    // Kasus B: Data berupa format panjang "Mon Feb 02..."
                    const dateObj = new Date(item.tanggal);
                    if (!isNaN(dateObj)) {
                        const idD = String(dateObj.getDate()).padStart(2, '0');
                        const idM = String(dateObj.getMonth() + 1).padStart(2, '0');
                        const idY = dateObj.getFullYear();
                        itemFormatted = `${idD}/${idM}/${idY}`;
                    }
                }
                
                return itemFormatted === targetDateStr;
            });

            dailyAgenda.forEach(item => {
                if (item.jam && item.jam.includes('-')) {
                    const parts = item.jam.split('-');
                    const mulai = parts[0].trim();
                    const selesai = parts[1].trim();

                    const idxMulai = jamOperasional.indexOf(mulai);
                    const idxSelesai = jamOperasional.indexOf(selesai);

                    if (idxMulai !== -1 && idxSelesai !== -1) {
                        renderVisualAgenda(selectedDate, idxMulai, idxSelesai, item.ruang, item.acara, item.pic, item.emailPIC, item.orderId);
                    }
                }
            });

            renderCalendar();

        } catch (err) {
            console.error("Gagal sinkronisasi data:", err);
        } finally {
            hideTableLoading();
            if (Swal.isVisible()) Swal.close();
        }
    }

    // Fungsi loading tabel
    function showTableLoading() {
        tableLoading.classList.remove('hidden');
    }

    function hideTableLoading() {
        tableLoading.classList.add('hidden');
    }

    // Fungsi Sinkronisasi Data
    const btnSync = document.getElementById('btnSync');
    if (btnSync) {
        btnSync.onclick = () => {
            Swal.fire({
                title: 'Sinkronisasi Data...',
                text: 'Mengambil jadwal terbaru dari Cloud',
                allowOutsideClick: false,
                didOpen: () => {
                    Swal.showLoading();
                    renderDailyTable(); // Ini akan memicu loadAndRenderAgenda
                }
            });
        };
    }

    // Fungsi untuk filter "Pesanan Saya"
    const btnMyBooking = document.getElementById('btnMyBooking');
    if (btnMyBooking) {
        btnMyBooking.onclick = function() {
            if (!sessionStorage.getItem("user")) {
                return Swal.fire('Akses Terbatas', 'Silakan Login untuk melihat pesanan Anda.', 'warning');
            }
            isMyBookingFilterActive = !isMyBookingFilterActive;
            
            // Efek Visual Tombol Aktif
            this.style.background = isMyBookingFilterActive ? "#2b6cb0" : "#edf2f7";
            this.style.color = isMyBookingFilterActive ? "white" : "#2d3748";
            
            renderDailyTable();
        };
    }

    // Fungsi untuk filter "Ruangan Tertentu"
    const btnFilter = document.getElementById('btnFilterRuang');
    if (btnFilter) {
        btnFilter.onclick = async () => {
            const { value: selectedRuangan } = await Swal.fire({
                title: 'Filter Ruangan',
                html: `
                    <div style="display: flex; justify-content: space-between; margin-bottom: 15px; border-bottom: 1px solid #eee; padding-bottom: 10px;">
                        <button type="button" class="swal2-confirm swal2-styled" style="margin:0; padding: 5px 10px; font-size: 0.8rem; background-color: #4a5568;" id="selectAll">Pilih Semua</button>
                        <button type="button" class="swal2-cancel swal2-styled" style="margin:0; padding: 5px 10px; font-size: 0.8rem; background-color: #a0aec0;" id="deselectAll">Hapus Semua</button>
                    </div>
                    
                    <div id="filter-checkboxes" style="display: grid; grid-template-columns: 1fr 1fr; text-align: left; gap: 10px; font-size: 0.9rem; max-height: 300px; overflow-y: auto; padding: 5px;">
                        ${daftarRuang.map(r => `
                            <label style="cursor:pointer; display: flex; align-items: center; gap: 8px;">
                                <input type="checkbox" value="${r}" ${filterRuangAktif.includes(r) ? 'checked' : ''}> 
                                <span>${r}</span>
                            </label>
                        `).join('')}
                    </div>`,
                showCancelButton: true,
                confirmButtonText: 'Terapkan Filter',
                didOpen: () => {
                    // Logika Pilih Semua
                    document.getElementById('selectAll').onclick = () => {
                        const checkboxes = document.querySelectorAll('#filter-checkboxes input[type="checkbox"]');
                        checkboxes.forEach(cb => cb.checked = true);
                    };
                    // Logika Hapus Semua
                    document.getElementById('deselectAll').onclick = () => {
                        const checkboxes = document.querySelectorAll('#filter-checkboxes input[type="checkbox"]');
                        checkboxes.forEach(cb => cb.checked = false);
                    };
                },
                preConfirm: () => {
                    const checked = Array.from(document.querySelectorAll('#filter-checkboxes input:checked')).map(el => el.value);
                    if (checked.length === 0) {
                        Swal.showValidationMessage('Pilih minimal satu ruangan!');
                    }
                    return checked;
                }
            });

            if (selectedRuangan) {
                filterRuangAktif = selectedRuangan;
                renderDailyTable();
            }
        };
    }

    // Fungsi untuk menunjukkan detail modal
    async function showDetailModal(data) {
        const user = JSON.parse(sessionStorage.getItem("user") || "{}");
        const userEmail = user.email ? user.email.toLowerCase().trim() : "";
        
        // Logika Hak Akses: Pemilik OR Admin
        const isOwner = userEmail === (data.emailPIC ? data.emailPIC.toLowerCase().trim() : "");
        const isAdmin = ADMIN_LIST.includes(userEmail);
        const hasAccess = isOwner || isAdmin;

        await Swal.fire({
            title: '', // Kita kosongkan judul bawaan agar bisa kita desain kustom di HTML
            html: `
                <div style="text-align: left;">
                    <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px; border-bottom: 1px solid #edf2f7; padding-bottom: 15px;">
                        <div style="font-size: 1.25rem; font-weight: 700; color: #1a3a5f; padding-right: 10px;">
                            Detil Informasi Penggunaan Ruang
                        </div>
                        ${hasAccess ? `
                            <div style="display: flex; gap: 6px; flex-shrink: 0;">
                                <button id="editBtn" class="btn-icon-detail" title="Edit">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7M18.5 2.5a2.121 2.121 0 113 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>
                                </button>
                                <button id="deleteBtn" class="btn-icon-detail btn-danger-detail" title="Hapus">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                                </button>
                            </div>
                        ` : ''}
                    </div>

                    <div class="form-group">
                        <label>Tanggal Penggunaan</label>
                        <input type="text" value="${data.tanggal}" readonly style="background: #f8fafc; color: #718096;">
                    </div>
                    <div class="form-group">
                        <label>Ruangan</label>
                        <input type="text" value="${data.ruang}" readonly style="background: #f8fafc; color: #718096;">
                    </div>
                    <div class="form-group">
                        <label>Waktu Kegiatan</label>
                        <input type="text" value="${data.jam}" readonly style="background: #f8fafc; color: #718096;">
                    </div>
                    <div class="form-group">
                        <label>Peruntukan Acara</label>
                        <textarea id="detAcara" readonly style="border: 1px solid #e2e8f0;">${data.acara}</textarea>
                    </div>
                    <div class="form-group">
                        <label>PIC Kegiatan</label>
                        <input type="text" id="detPIC" value="${data.pic}" readonly style="border: 1px solid #e2e8f0;">
                    </div>

                    <div id="saveContainer" style="display:none; margin-top: 20px;">
                        <button id="saveUpdateBtn" class="btn-submit" style="background: #38a169;">Simpan Perubahan Agenda</button>
                    </div>
                </div>
            `,
            showConfirmButton: false,
            showCloseButton: true,
            customClass: { popup: 'swal-professional-popup' },
            didOpen: () => {
                if (hasAccess) {
                    const editBtn = document.getElementById('editBtn');
                    const deleteBtn = document.getElementById('deleteBtn');
                    const saveContainer = document.getElementById('saveContainer');

                    editBtn.onclick = () => {
                        document.getElementById('detAcara').readOnly = false;
                        document.getElementById('detPIC').readOnly = false;
                        document.getElementById('detAcara').style.borderColor = "#3182ce";
                        document.getElementById('detPIC').style.borderColor = "#3182ce";
                        saveContainer.style.display = "block";
                        editBtn.style.opacity = "0.5";
                        editBtn.disabled = true;
                    };

                    document.getElementById('saveUpdateBtn').onclick = () => processUpdate(data);
                    deleteBtn.onclick = () => processDelete(data.orderId);
                }
            }
        });
    }

    // Fungsi untuk edit dan hapus agenda
    function enableEditMode(data) {
        // Ubah input jadi bisa diketik
        document.getElementById('detAcara').readOnly = false;
        document.getElementById('detPIC').readOnly = false;
        
        // Tambahkan tombol simpan di bawah
        const container = document.querySelector('.swal2-html-container');
        const saveBtn = document.createElement('button');
        saveBtn.innerText = "Simpan Perubahan Agenda";
        saveBtn.className = "btn-submit";
        saveBtn.style.marginTop = "20px";
        saveBtn.onclick = () => processUpdate(data);
        container.appendChild(saveBtn);
    }

    async function processUpdate(oldData) {
        // Ambil input terbaru dari modal
        const inputAcara = document.getElementById('detAcara').value;
        const inputPIC = document.getElementById('detPIC').value;

        // Hitung ulang "Hari" agar tidak kosong di spreadsheet
        const parts = oldData.tanggal.split('/');
        const dObj = new Date(parts[2], parts[1] - 1, parts[0]);
        const hariIndo = dObj.toLocaleDateString('id-ID', { weekday: 'long' });

        const payload = {
            "action": "updateBooking",
            "Order ID": oldData.orderId,
            "Hari": hariIndo, // Pastikan Hari disertakan kembali
            "Tanggal": oldData.tanggal,
            "Acara": inputAcara,
            "Jam": oldData.jam, // Gunakan jam asli agar tidak berubah
            "PIC Kegiatan": inputPIC,
            "ruang": oldData.ruang,
            "Email PIC": oldData.emailPIC
        };

        Swal.fire({ title: 'Menyimpan...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
        
        try {
            const fd = new FormData();
            fd.append("data", JSON.stringify(payload));

            const res = await fetch("https://script.google.com/macros/s/AKfycbxQ8poOfZ9y4dZVEOQ0qxWxBJ9i8P0-HlHkhCEBYWpooP7kHobijeUcHc6uHCc6uyN5/exec", { method: "POST", body: fd });
            const result = await res.json();

            if (result.result === "success") {
            Swal.fire('Berhasil!', 'Agenda telah diperbarui.', 'success').then(() => renderDailyTable());
            } else {
                throw new Error(result.message);
            }
        } catch (err) {
            Swal.fire('Gagal Update', 'Terjadi kesalahan koneksi atau CORS.', 'error');
        }
    }   

    async function processDelete(orderId) {
        const confirm = await Swal.fire({
            title: 'Hapus Agenda?',
            text: "Tindakan ini tidak dapat dibatalkan!",
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#d33',
            confirmButtonText: 'Ya, Hapus!'
        });

        if (confirm.isConfirmed) {
            Swal.fire({ title: 'Menghapus...', didOpen: () => Swal.showLoading() });
            const fd = new FormData();
            fd.append("data", JSON.stringify({ action: "deleteBooking", "Order ID": orderId }));
            
            await fetch("https://script.google.com/macros/s/AKfycbxQ8poOfZ9y4dZVEOQ0qxWxBJ9i8P0-HlHkhCEBYWpooP7kHobijeUcHc6uHCc6uyN5/exec", { method: "POST", body: fd });
            Swal.fire('Terhapus!', 'Agenda telah dihapus.', 'success').then(() => renderDailyTable());
        }
    }

    // Fungsi untuk view per minggu:
    document.getElementById('btnToggleView').onclick = function() {
        const textSpan = document.getElementById('textToggleView');
        
        if (currentViewMode === 'daily') {
            currentViewMode = 'weekly';
            textSpan.innerText = "View per Hari";
            renderWeeklyTable(); // Fungsi baru
        } else {
            currentViewMode = 'daily';
            textSpan.innerText = "View per Minggu";
            renderDailyTable(); // Kembali ke harian
        }
    };

    // Fungsi untuk Render Tabel Mingguan
    function renderWeeklyTable() {
        const tableHead = document.querySelector('#agendaTable thead tr');
        const tableBody = document.getElementById('tableBody');
        
        // 1. Bersihkan tabel
        tableHead.innerHTML = '<th class="time-column">Hari</th>';
        tableBody.innerHTML = '';

        // 2. Buat Header Ruangan (Sesuai filterRuangAktif)
        filterRuangAktif.forEach((ruang) => {
            const th = document.createElement('th');
            th.innerText = ruang;
            tableHead.appendChild(th);
        });

        // 3. Tentukan Range Minggu Ini (Senin - Minggu)
        const curr = new Date(selectedDate);
        const day = curr.getDay();
        const diff = curr.getDate() - day + (day === 0 ? -6 : 1); // Penyesuaian ke hari Senin
        const monday = new Date(curr.setDate(diff));
        
        const namaHari = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"];

        // 4. Buat Baris per Hari dalam 1 Minggu
        for (let i = 0; i < 7; i++) {
            const rowDate = new Date(monday);
            rowDate.setDate(monday.getDate() + i);
            
            const row = document.createElement('tr');
            
            // Kolom Nama Hari & Tanggal
            let dateLabel = `${String(rowDate.getDate()).padStart(2, '0')}/${String(rowDate.getMonth() + 1).padStart(2, '0')}`;
            row.innerHTML = `<td class="time-column">
                <div class="day-name">${namaHari[i]}</div>
                <div class="hour-label">${dateLabel}</div>
            </td>`;

            // Kolom Ruangan
            filterRuangAktif.forEach((ruang) => {
                const ruangIdxAsli = daftarRuang.indexOf(ruang);
                const cellId = `weekly-${rowDate.getFullYear()}-${rowDate.getMonth()}-${rowDate.getDate()}-${ruangIdxAsli}`;
                row.innerHTML += `<td id="${cellId}" style="height:100px; vertical-align:top; padding:8px; background: white;"></td>`;
            });

            tableBody.appendChild(row);
        }

        // 5. Isi data
        fillWeeklyData(monday);
    }

    function fillWeeklyData(mondayDate) {
        const sundayDate = new Date(mondayDate);
        sundayDate.setDate(mondayDate.getDate() + 6);

        // Object untuk menghitung jumlah agenda per sel
        const cellCounts = {}; 

        // Langkah 1: Hitung jumlah agenda per sel
        allBookedData.forEach(item => {
            let itemDate;
            if (item.tanggal.includes('/')) {
                const p = item.tanggal.split('/');
                itemDate = new Date(p[2], p[1] - 1, p[0]);
            } else {
                itemDate = new Date(item.tanggal);
            }

            itemDate.setHours(0,0,0,0);
            const start = new Date(mondayDate); start.setHours(0,0,0,0);
            const end = new Date(sundayDate); end.setHours(0,0,0,0);

            if (itemDate >= start && itemDate <= end) {
                const ruangIdxAsli = daftarRuang.indexOf(item.ruang);
                const cellId = `weekly-${itemDate.getFullYear()}-${itemDate.getMonth()}-${itemDate.getDate()}-${ruangIdxAsli}`;
                
                cellCounts[cellId] = (cellCounts[cellId] || 0) + 1;
            }
        });

        // Langkah 2: Render satu blok per sel berdasarkan hitungan
        for (const cellId in cellCounts) {
            const cell = document.getElementById(cellId);
            if (cell) {
                const count = cellCounts[cellId];
                const block = document.createElement('div');
                block.className = 'agenda-block agenda-start agenda-end';
                
                // Warna makin gelap jika makin penuh (opsional)
                block.style.backgroundColor = count > 3 ? '#1a3a5f' : '#2b6cb0'; 
                
                block.style.margin = 'auto';
                block.style.width = '90%';
                block.style.height = '40px';
                block.style.display = 'flex';
                block.style.alignItems = 'center';
                block.style.justifyContent = 'center';
                block.style.borderRadius = '6px';
                block.style.cursor = 'pointer';
                block.style.color = 'white';
                
                block.innerHTML = `<span style="font-size: 0.7rem;">TERISI (${count})</span>`;
                
                block.onclick = () => {
                    // Ekstrak tanggal dari ID sel (format: weekly-YYYY-MM-DD-Idx)
                    const parts = cellId.split('-');
                    const clickedDate = new Date(parts[1], parts[2], parts[3]);
                    
                    // Pindah ke harian
                    window.selectDate(clickedDate.getDate());
                    if (currentViewMode === 'weekly') {
                        document.getElementById('btnToggleView').click();
                    }
                };
                
                cell.appendChild(block);
            }
        }
    }
});