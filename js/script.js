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
    let listDateAnchor = new Date(selectedDate);

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
    const daftarRuang = [
        { kode: "RS.1", nama: "RS.1", kapasitas: 44 },
        { kode: "RS.2", nama: "RS.2", kapasitas: 26 },
        { kode: "RS.3", nama: "RS.3", kapasitas: 12 },
        { kode: "Lab. Hidro", nama: "Lab. Hidro", kapasitas: 8 },
        { kode: "Lab GGGF", nama: "Lab GGGF", kapasitas: 16 },
        { kode: "Lab CAGE", nama: "Lab CAGE", kapasitas: 16 },
        { kode: "Lab Foto", nama: "Lab Foto", kapasitas: 70 },
        { kode: "Lab Geokom", nama: "Lab Geokom", kapasitas: 75 },
        { kode: "III.6", nama: "III.6", kapasitas: 12 },
        { kode: "III.5", nama: "III.5", kapasitas: 45 },
        { kode: "III.4", nama: "III.4", kapasitas: 190 },
        { kode: "III.3", nama: "III.3", kapasitas: 80 },
        { kode: "III.2", nama: "III.2", kapasitas: 80 },
        { kode: "III.1", nama: "III.1", kapasitas: 117 },
        { kode: "1.1", nama: "1.1", kapasitas: 80 },
        { kode: "R. Pengurus", nama: "R. Pengurus", kapasitas: 8 },
        { kode: "R. Sidang SURTA", nama: "R. Sidang SURTA", kapasitas: 8 },
        { kode: "R. Bersama", nama: "R. Bersama", kapasitas: 12 }
    ];

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
            daftarRuang.forEach(ruangObj => {
                selectRuang.innerHTML += `<option value="${ruangObj.kode}">${ruangObj.nama} (Kapasitas: ${ruangObj.kapasitas})</option>`;
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
        fetchHolidays();
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

        let currentMonthHolidays = [];
        let periodikSummary = [];

        for (let d = 1; d <= daysInMonth; d++) {
            // --- LOGIKA PENANDA TITIK ---
            const dStr = String(d).padStart(2, '0');
            const mStr = String(month + 1).padStart(2, '0');
            const dateKeyBooking = `${dStr}/${mStr}/${year}`;
            const dateKeyHoliday = `${year}-${mStr}-${dStr}`;

            // Hari Libur
            const holidayInfo = holidaysData[dateKeyHoliday];
            const isHoliday = holidayInfo && (holidayInfo.holiday === true || holidayInfo.summary);

            if (isHoliday) {
                const holidayName = holidayInfo.summary[0];
                currentMonthHolidays.push(`${d} ${monthNames[month]}: ${holidayName}`);
            }

            // Periodik
            const dayBookings = allBookedData.filter(item =>
                normalizeDateKey(item.tanggal) === dateKeyBooking
            );

            const hasBooking = dayBookings.length > 0;
            const periodikBookings = dayBookings.filter(item =>
                item.tipePesanan === 'periodik'
            );

            let periodikHtml = "";
            periodikBookings.forEach(book => {
                const baseId = book.orderId.split('-').slice(0, 2).join('-');
                const yesterdayKey = `${String(d-1).padStart(2,'0')}/${mStr}/${year}`;
                const tomorrowKey  = `${String(d+1).padStart(2,'0')}/${mStr}/${year}`;
                const hasYesterday = allBookedData.some(item =>
                    item.orderId.startsWith(baseId) &&
                    normalizeDateKey(item.tanggal) === yesterdayKey
                );
                const hasTomorrow = allBookedData.some(item =>
                    item.orderId.startsWith(baseId) &&
                    normalizeDateKey(item.tanggal) === tomorrowKey
                );
                const isStart = !hasYesterday;
                const isEnd   = !hasTomorrow;
                periodikHtml += `
                    <div class="periodik-line 
                        ${isStart ? 'periodik-start' : ''} 
                        ${isEnd ? 'periodik-end' : ''}"
                        title="${book.acara}">
                    </div>
                `;
                if (isStart) {
                    periodikSummary.push(
                        `${book.acara} (${book.ruang}) ${d} ${monthNames[month]}`
                    );
                }
            });

            const isSelected = (d === selectedDate.getDate() && month === selectedDate.getMonth() && year === selectedDate.getFullYear());
            const isToday = (d === new Date().getDate() && month === new Date().getMonth() && year === new Date().getFullYear());

            html += `
            <div class="cal-day ${isSelected ? 'cal-selected' : ''} ${isToday ? 'cal-today' : ''} ${isHoliday ? 'cal-holiday' : ''}" 
                onclick="selectDate(${d})"
                title="${isHoliday ? holidayInfo.summary[0] : ''}">
                ${periodikHtml}
                ${d}
                ${hasBooking ? '<span class="cal-dot"></span>' : ''}
            </div>`;
        }
        html += `</div>`;
        calendarUI.innerHTML = html;

        // Render daftar hari libur
        const listElem = document.getElementById('holiday-list');
        if (currentMonthHolidays.length > 0) {
            listElem.innerHTML = currentMonthHolidays.map(txt => `<li>${txt}</li>`).join('');
            document.getElementById('holiday-info').style.display = 'block';
        } else {
            document.getElementById('holiday-info').style.display = 'none';
        }

        // Render periodik
        const periodikElem = document.getElementById('periodik-list');
        const periodikContainer = document.getElementById('periodik-info');
        if (periodikSummary.length > 0) {
            periodikContainer.style.display = 'block';
            // Menghilangkan duplikasi nama acara yang sama di list bawah
            const uniquePeriodik = [...new Set(periodikSummary)];
            periodikElem.innerHTML = uniquePeriodik.map(txt => `<li>${txt}</li>`).join('');
        } else {
            periodikContainer.style.display = 'none';
        }
    }

    // Fungsi normalisasi tanggal
    function normalizeDateKey(tgl) {
        if (typeof tgl === 'string' && tgl.includes('/')) {
            return tgl;
        }
        const d = new Date(tgl);
        if (isNaN(d)) return '';
        return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
    }


    // Fungsi hari libur di kalender
    let holidaysData = {};
    async function fetchHolidays() {
        try {
            const response = await fetch('https://raw.githubusercontent.com/guangrei/APIHariLibur_V2/main/calendar.min.json');
            holidaysData = await response.json();
            renderCalendar(); // Gambar ulang kalender setelah data didapat
        } catch (error) {
            console.error("Gagal mengambil data hari libur:", error);
        }
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
        headerRow.innerHTML = `<th>Jam</th>` + filterRuangAktif.map(r => `<th>${r.nama}</th>`).join('');

        jamOperasional.forEach(jam => {
            let row = document.createElement('tr');
            let cells = `<td style="font-weight:bold; background:#f1f5f9;">${jam}</td>`;
            
            filterRuangAktif.forEach(r => {
                const idxAsli = daftarRuang.findIndex(ruang => ruang.kode === r.kode); // Mencari index asli untuk ID sel
                const id = `cell-${selectedDate.getFullYear()}-${selectedDate.getMonth()}-${selectedDate.getDate()}-${jam.replace(':','')}-${idxAsli}`;
                cells += `<td id="${id}"></td>`;
            });

            row.innerHTML = cells;
            tableBody.appendChild(row);
        });

        loadAndRenderAgenda();
        setTimeout(updateTimeMarker, 100);
    }

    // Fungsi ubah ke tanggal Indonesia
    function formatTanggalIndonesia(dateObj) {
        return dateObj.toLocaleDateString('id-ID', {
            weekday: 'long',
            day: 'numeric',
            month: 'long',
            year: 'numeric'
        });
    }

    // Logika Form Submit
    form.onsubmit = async (e) => {
        e.preventDefault();

        // 1. Persiapan Data
        const user = JSON.parse(sessionStorage.getItem("user") || "{}");
        const userEmail = user.email ? user.email.toLowerCase().trim() : "";
        const isAdmin = ADMIN_LIST.includes(userEmail);
        const mode = document.getElementById('bookingMode').value;
        const tglValue = document.getElementById('tglKegiatan').value;
        const tglInput = new Date(tglValue);
        tglInput.setHours(0,0,0,0);

        const today = new Date();
        today.setHours(0,0,0,0);

        // Ambil detail untuk cek tabrakan
        const ruang = selectRuang.value;
        const ruangIdx = daftarRuang.indexOf(ruang);
        const mulai = selectJamMulai.value;
        const selesai = selectJamSelesai.value;
        const idxMulai = jamOperasional.indexOf(mulai);
        const idxSelesai = jamOperasional.indexOf(selesai);
        const jumlahPartisipan = document.getElementById('jumlahPartisipan').value;

        // 2. Validasi Jam Dasar (Tetap sama)
        if (idxSelesai <= idxMulai) return Swal.fire('Error', 'Jam tidak valid!', 'error');

        // --- 2a. LOGIKA VALIDASI DURASI & H-3 (KHUSUS NON-ADMIN) ---
        if (!isAdmin) {
            const durasiJam = idxSelesai - idxMulai;
            
            // Hitung selisih hari (H-x)
            const diffTime = tglInput - today;
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

            if (durasiJam >= 5) {
                let datesToCheck = [];
                if (mode === 'harian') {
                    datesToCheck = [tglInput];
                } else {
                    const startDate = new Date(document.getElementById('tglAwal').value);
                    const endDate   = new Date(document.getElementById('tglAkhir').value);
                    startDate.setHours(0,0,0,0);
                    endDate.setHours(0,0,0,0);
                    if (mode === 'periodik') {
                        // Semua hari dalam range
                        datesToCheck = getDatesInRange(startDate, endDate);
                    } else if (mode === 'berulang') {
                        const selectedDays = Array.from(
                            document.querySelectorAll('input[name="repeatDay"]:checked')
                        ).map(cb => parseInt(cb.value)); // 0–6
                        datesToCheck = getDatesInRange(startDate, endDate, selectedDays);
                    }
                }

                for (const cursor of datesToCheck) {
                    const diffTime = cursor - today;
                    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                    if (diffDays < 3) {
                        return Swal.fire({
                            icon: 'warning',
                            title: 'Pemesanan Ditolak',
                            text: 'Pemesanan ruang dengan waktu yang lama (5 jam atau lebih) hanya dapat dilakukan H-3 Acara. Jika ada acara mendadak atau terdesak pada hari H, silahkan lakukan pemesanan via Admin (Hanifah).',
                            confirmButtonColor: '#3085d6'
                        });
                    }
                }
            }
        }

        // --- 2b. LOGIKA CEK TABRAKAN (CONFLICT CHECK) ---
        const y = tglInput.getFullYear();
        const m = tglInput.getMonth();
        const d = tglInput.getDate();
        
        const isConflict = hasConflictData(tglInput, ruang, idxMulai, idxSelesai);

        if (isConflict) {
            return Swal.fire({
                icon: 'error',
                title: 'Jadwal Bentrok!',
                text: 'Anda tidak bisa memesan ruangan di jam tersebut karena sudah ada agenda lainnya.',
                confirmButtonColor: '#e53e3e'
            });
        }

        // --- 2c. LOGIKA VALIDASI KAPASITAS RUANG ---
        if (!validateRoomCapacity(ruang, jumlahPartisipan, isAdmin)) {
            return;
        }
        
        // 3. Konfirmasi & Loading (Baru dijalankan jika tidak bentrok)
        const payload = {
            "Order ID": "DTGD-" + Date.now(),
            "Mode": mode,
            "Tipe Pesanan": mode,
            "Acara": document.getElementById('namaAcara').value,
            "Jam": `${selectJamMulai.value} - ${selectJamSelesai.value}`,
            "PIC Kegiatan": document.getElementById('picKegiatan').value,
            "Jumlah Partisipan": document.getElementById('jumlahPartisipan').value || "",
            "Email PIC": user.email || "",
            "ruang": selectRuang.value
        };

        if (mode === 'harian') {
            const tgl = new Date(document.getElementById('tglKegiatan').value);
            payload["Tanggal"] = tgl.toLocaleDateString('id-ID', { day: '2-digit', month: '2-digit', year: 'numeric' });
            payload["Hari"] = tgl.toLocaleDateString('id-ID', { weekday: 'long' });
        } else {
            payload["Tanggal Awal"] = document.getElementById('tglAwal').value;
            payload["Tanggal Akhir"] = document.getElementById('tglAkhir').value;
            
            if (mode === 'berulang') {
                const selectedDays = Array.from(document.querySelectorAll('input[name="repeatDay"]:checked')).map(cb => cb.value);
                payload["Hari Perulangan"] = selectedDays.join(','); // Contoh: "1,3,5" (Sen, Rab, Jum)
            }
        }

        if (mode !== 'harian') {
            const startDate = new Date(payload["Tanggal Awal"]);
            const endDate   = new Date(payload["Tanggal Akhir"]);
            startDate.setHours(0,0,0,0);
            endDate.setHours(0,0,0,0);
            const repeatDays = payload["Hari Perulangan"]
                ? payload["Hari Perulangan"].split(',').map(Number)
                : [];
            let cursor = new Date(startDate);
            while (cursor <= endDate) {
                const day = cursor.getDay(); // 0-6
                if (mode === 'periodik' || repeatDays.includes(day)) {
                    if (hasConflictData(cursor, ruang, idxMulai, idxSelesai)) {
                        return Swal.fire({
                            icon: 'error',
                            title: 'Jadwal Bentrok!',
                            text: `Bentrok pada hari ${formatTanggalIndonesia(cursor)}\n
                            Ruang ${ruang}, pukul ${mulai} – ${selesai}`
                        });
                    }
                }
                cursor.setDate(cursor.getDate() + 1);
            }
        }
        
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

                const response = await fetch("https://script.google.com/macros/s/AKfycbxdighuUGoWrArL6JREQBrp4ikns0fyZDpFQ-kxHrp_Hj9tRrcD3BkAv4XA_CoOSaIH/exec", {
                    method: "POST",
                    body: fd // Kirim sebagai FormData
                });

                const result = await response.json();

                if (result.result === "success") {
                    await loadAndRenderAgenda();

                    renderCalendar();
                    renderDailyTable();
                    
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

    // Fungsi konflik
    function isTimeOverlap(startA, endA, startB, endB) {
        return startA < endB && endA > startB;
    }

    function hasConflictData(targetDate, ruang, idxMulai, idxSelesai) {
        return allBookedData.some(item => {

            // Cek ruang
            if (item.ruang !== ruang) return false;

            // Normalisasi tanggal item
            let itemDate;
            if (typeof item.tanggal === 'string' && item.tanggal.includes('/')) {
                const [d,m,y] = item.tanggal.split('/');
                itemDate = new Date(y, m-1, d);
            } else {
                itemDate = new Date(item.tanggal);
            }
            itemDate.setHours(0,0,0,0);

            if (itemDate.getTime() !== targetDate.getTime()) return false;

            // Ambil jam item
            const [jamMulaiItem, jamSelesaiItem] = item.jam.split(' - ');
            const idxMulaiItem = jamOperasional.indexOf(jamMulaiItem);
            const idxSelesaiItem = jamOperasional.indexOf(jamSelesaiItem);

            return isTimeOverlap(
                idxMulai, idxSelesai,
                idxMulaiItem, idxSelesaiItem
            );
        });
    }

    // Fungsi helper untuk menolak pemesanan lama di periodik
    function getDatesInRange(start, end, allowedDays = null) {
        const dates = [];
        const cursor = new Date(start);
        cursor.setHours(0,0,0,0);

        while (cursor <= end) {
            if (!allowedDays || allowedDays.includes(cursor.getDay())) {
                dates.push(new Date(cursor));
            }
            cursor.setDate(cursor.getDate() + 1);
        }
        return dates;
    }

    // Window untuk ganti tipe pemesanan ruang
    window.switchBookingTab = (mode) => {
        // 1. Update State
        document.getElementById('bookingMode').value = mode;
        
        // 2. Update Visual Tab
        document.querySelectorAll('.btn-tab').forEach(btn => {
            btn.classList.remove('active');
            if (btn.innerText.toLowerCase() === mode) btn.classList.add('active');
        });

        // 3. Kontrol Visibilitas Input
        const tunggal = document.getElementById('group-tanggal-tunggal');
        const range = document.getElementById('group-tanggal-range');
        const repeatDays = document.getElementById('group-hari-berulang');

        // Ambil input aslinya
        const inputTglTunggal = document.getElementById('tglKegiatan');
        const inputTglAwal = document.getElementById('tglAwal');
        const inputTglAkhir = document.getElementById('tglAkhir');

        if (mode === 'harian') {
            tunggal.style.display = 'block';
            range.style.display = 'none';
            repeatDays.style.display = 'none';
            inputTglTunggal.setAttribute('required', '');
            inputTglAwal.removeAttribute('required');
            inputTglAkhir.removeAttribute('required');
        } else {
            tunggal.style.display = 'none';
            range.style.display = 'block';
            
            // Matikan required harian, aktifkan range
            inputTglTunggal.removeAttribute('required');
            inputTglAwal.setAttribute('required', '');
            inputTglAkhir.setAttribute('required', '');

            if (mode === 'berulang') {
                repeatDays.style.display = 'block';
            } else {
                repeatDays.style.display = 'none';
            }
        }
    };

    // Fungsi untuk validasi kapasitas ruang
    function validateRoomCapacity(ruang, jumlahPartisipan, isAdmin = false) {
        // Admin boleh override (opsional, bisa dihapus kalau tidak mau)
        if (isAdmin) return true;

        const room = daftarRuang.findIndex(r => r.kode === ruang || r.nama === ruang);
        if (room === -1) return true; // fallback aman

        const kapasitas = daftarRuang[room].kapasitas;
        const jumlah = parseInt(jumlahPartisipan || 0);

        if (jumlah > kapasitas) {
            Swal.fire({
                icon: 'warning',
                title: 'Kapasitas Ruangan Terlampaui',
                html: `
                    Ruangan <b>${ruang}</b> memiliki kapasitas maksimal 
                    <b>${kapasitas} orang</b>.<br><br>
                    Jumlah partisipan yang diinput: 
                    <b>${jumlah} orang</b>.
                `,
                confirmButtonColor: '#e53e3e'
            });
            return false;
        }
        return true;
    }

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

        const y = dateObj.getFullYear();
        const m = dateObj.getMonth();
        const d = dateObj.getDate();
        const ruangIdx = daftarRuang.findIndex(room => room.kode === ruang || room.nama === ruang);

        // Logika Warna
        const palette = [
        '#3182ce', '#38a169', '#d69e2e', '#e53e3e', '#805ad5',
        '#319795', '#d53f8c', '#4a5568', '#2b6cb0', '#2c7a7b',
        '#cd8338', '#5a67d8', '#97266d', '#285e61', '#c05621'
        ];

        // Logika agar warna konsisten berdasarkan Order ID (tidak berubah saat refresh)
        const getConsistentColor = (id) => {
            let hash = 0;
            const str = id.toString();
            for (let i = 0; i < str.length; i++) {
                hash = str.charCodeAt(i) + ((hash << 5) - hash);
            }
            return palette[Math.abs(hash) % palette.length];
        };

        let finalColor = getConsistentColor(orderID);
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
            const response = await fetch("https://script.google.com/macros/s/AKfycbxdighuUGoWrArL6JREQBrp4ikns0fyZDpFQ-kxHrp_Hj9tRrcD3BkAv4XA_CoOSaIH/exec?t=" + Date.now());
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
            const { value: selectedKodes } = await Swal.fire({
                title: 'Filter Ruangan',
                html: `
                    <div style="display: flex; justify-content: space-between; margin-bottom: 15px; border-bottom: 1px solid #eee; padding-bottom: 10px;">
                        <button type="button" class="swal2-confirm swal2-styled" style="margin:0; padding: 5px 10px; font-size: 0.8rem; background-color: #4a5568;" id="selectAll">Pilih Semua</button>
                        <button type="button" class="swal2-cancel swal2-styled" style="margin:0; padding: 5px 10px; font-size: 0.8rem; background-color: #a0aec0;" id="deselectAll">Hapus Semua</button>
                    </div>
                    
                    <div id="filter-checkboxes" style="display: grid; grid-template-columns: 1fr 1fr; text-align: left; gap: 10px; font-size: 0.9rem; max-height: 300px; overflow-y: auto; padding: 5px;">
                        ${daftarRuang.map(r => {
                            const isChecked = filterRuangAktif.some(akt => akt.kode === r.kode);
                            return `
                                <label style="cursor:pointer; display: flex; align-items: center; gap: 8px;">
                                    <input type="checkbox" value="${r.kode}" ${isChecked ? 'checked' : ''}> 
                                    <span>${r.nama}</span>
                                </label>
                            `;
                        }).join('')}
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
                    const checkedKodes = Array.from(document.querySelectorAll('#filter-checkboxes input:checked')).map(el => el.value);
                    if (checkedKodes.length === 0) {
                        Swal.showValidationMessage('Pilih minimal satu ruangan!');
                    }
                    return checkedKodes;
                }
            });

            if (selectedKodes) {
                filterRuangAktif = daftarRuang.filter(r => selectedKodes.includes(r.kode));
                if (currentViewMode === 'weekly') {
                    renderWeeklyTable();
                } else {
                    renderDailyTable();
                }
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

        let tanggalFormatted = data.tanggal;
        try {
            if (typeof data.tanggal === 'string' && data.tanggal.includes('/')) {
                const p = data.tanggal.split('/');
                const dateObj = new Date(p[2], p[1] - 1, p[0]);
                tanggalFormatted = dateObj.toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
            } else {
                const dateObj = new Date(data.tanggal);
                tanggalFormatted = dateObj.toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
            }
        } catch (e) {
            console.error("Gagal format tanggal:", e);
        }

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
                        <input type="text" value="${tanggalFormatted}" readonly style="background: #f8fafc; color: #718096;">
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
                        <label>Jumlah Partisipan</label>
                        <input type="number" id="detPeserta" value="${data.peserta || ''}" readonly style="border: 1px solid #e2e8f0">
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
                        const fields = ['detAcara', 'detPIC', 'detPeserta'];
                        fields.forEach(id => {
                            const el = document.getElementById(id);
                            el.readOnly = false;
                            el.style.borderColor = "#3182ce";
                            el.style.background = "#fff";
                            el.style.color = "#000";
                        });

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
        const inputPeserta = document.getElementById('detPeserta').value;

        // Hitung ulang "Hari" agar tidak kosong di spreadsheet
        let dObj;
        // Cek apakah formatnya DD/MM/YYYY
        if (typeof oldData.tanggal === 'string' && oldData.tanggal.includes('/')) {
            const parts = oldData.tanggal.split('/');
            dObj = new Date(parts[2], parts[1] - 1, parts[0]);
        } else {
            // Jika formatnya objek Date atau string panjang (Thu Feb 05...)
            dObj = new Date(oldData.tanggal);
        }

        if (isNaN(dObj.getTime())) {
            return Swal.fire('Error', 'Data tanggal tidak terbaca dengan benar. Silakan coba refresh halaman.', 'error');
        }

        const hariIndo = dObj.toLocaleDateString('id-ID', { weekday: 'long' });
        const tglBersih = dObj.toLocaleDateString('id-ID', { day: '2-digit', month: '2-digit', year: 'numeric' });

        const payload = {
            "action": "updateBooking",
            "Order ID": oldData.orderId,
            "Hari": hariIndo, // Pastikan Hari disertakan kembali
            "Tanggal": tglBersih,
            "Acara": inputAcara,
            "Jam": oldData.jam, // Gunakan jam asli agar tidak berubah
            "PIC Kegiatan": inputPIC,
            "Jumlah Partisipan": inputPeserta || "",
            "ruang": oldData.ruang,
            "Email PIC": oldData.emailPIC
        };

        Swal.fire({ title: 'Menyimpan...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
        
        try {
            const fd = new FormData();
            fd.append("data", JSON.stringify(payload));

            const res = await fetch("https://script.google.com/macros/s/AKfycbxdighuUGoWrArL6JREQBrp4ikns0fyZDpFQ-kxHrp_Hj9tRrcD3BkAv4XA_CoOSaIH/exec", { method: "POST", body: fd });
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
            
            await fetch("https://script.google.com/macros/s/AKfycbxdighuUGoWrArL6JREQBrp4ikns0fyZDpFQ-kxHrp_Hj9tRrcD3BkAv4XA_CoOSaIH/exec", { method: "POST", body: fd });
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

        updateTimeMarker();
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
            th.innerText = ruang.nama;
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
            row.classList.add('weekly-row');
            
            // Kolom Nama Hari & Tanggal
            let dateLabel = `${String(rowDate.getDate()).padStart(2, '0')}/${String(rowDate.getMonth() + 1).padStart(2, '0')}`;
            row.innerHTML = `<td class="time-column">
                <div class="day-name">${namaHari[i]}</div>
                <div class="hour-label">${dateLabel}</div>
            </td>`;

            // Kolom Ruangan
            filterRuangAktif.forEach((ruang) => {
                const ruangIdxAsli = daftarRuang.findIndex(room => room.kode === ruang.kode);
                const cellId = `weekly-${rowDate.getFullYear()}-${rowDate.getMonth()}-${rowDate.getDate()}-${ruangIdxAsli}`;
                row.innerHTML += `<td id="${cellId}" class="weekly-cell"></td>`;
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
                const ruangIdxAsli = daftarRuang.findIndex(room => room.kode === item.ruang || room.nama === item.ruang);
                const cellId = `weekly-${itemDate.getFullYear()}-${itemDate.getMonth()}-${itemDate.getDate()}-${ruangIdxAsli}`;
                
                cellCounts[cellId] = (cellCounts[cellId] || 0) + 1;
            }
        });

        document.querySelectorAll('.weekly-cell').forEach(c => c.innerHTML = '');

        // Langkah 2: Render satu blok per sel berdasarkan hitungan
        for (const cellId in cellCounts) {
            const cell = document.getElementById(cellId);
            if (cell) {
                const count = cellCounts[cellId];
                const block = document.createElement('div');
                block.className = 'agenda-block agenda-start agenda-end';
                
                // Warna makin gelap jika makin penuh (opsional)
                block.style.backgroundColor = count > 3 ? '#1a3a5f' : '#2b6cb0'; 
                
                block.innerHTML = `<span>TERISI (${count})</span>`;
                
                block.onclick = () => {
                    const parts = cellId.split('-');
                    const d = parseInt(parts[3]);
                    window.selectDate(d);
                    document.getElementById('btnToggleView').click();
                };
                
                cell.appendChild(block);
            }
        }
    }

    // Fungsi garis marker waktu saat ini
    function updateTimeMarker() {
        // 1. CARI DAN HAPUS marker lama di awal fungsi (Kunci agar garis tidak nyangkut)
        const oldMarker = document.getElementById('time-marker');
        if (oldMarker) oldMarker.remove();

        // 2. VALIDASI: Jika View Minggu ATAU bukan tanggal hari ini, langsung keluar
        const today = new Date();
        if (currentViewMode === 'weekly' || selectedDate.toDateString() !== today.toDateString()) {
            return;
        }

        const now = today.getHours() + today.getMinutes() / 60;
        const startHour = parseInt(jamOperasional[0].split(':')[0]);
        const endHour = parseInt(jamOperasional[jamOperasional.length - 1].split(':')[0]) + 1;

        if (now < startHour || now > endHour) return;

        // 3. Cari baris berdasarkan jam sekarang
        const rows = document.querySelectorAll('#tableBody tr');
        const hourIndex = Math.floor(now - startHour);
        const targetRow = rows[hourIndex];

        if (targetRow) {
            const marker = document.createElement('div');
            marker.id = 'time-marker';
            marker.className = 'time-marker';
            
            // 4. Kalkulasi Posisi Vertikal (Y)
            const rowRect = targetRow.getBoundingClientRect();
            const wrapper = document.querySelector('.table-wrapper');
            const wrapperRect = wrapper.getBoundingClientRect();
            
            const rowHeight = targetRow.offsetHeight;
            const minuteOffset = (today.getMinutes() / 60) * rowHeight;
            
            // Jarak dari atas wrapper ke baris + offset menit
            const topPosition = (rowRect.top - wrapperRect.top) + wrapper.scrollTop + minuteOffset;
            marker.style.top = `${topPosition}px`;
            
            // 5. PERBAIKAN LEBAR: Gunakan scrollWidth agar garis memanjang sampai ujung kanan tabel
            const tableElement = document.getElementById('agendaTable');
            marker.style.width = `${tableElement.scrollWidth}px`; 
            
            wrapper.appendChild(marker);
        }
    }

    // Jalankan setiap 1 menit agar garis bergerak real-time
    setInterval(updateTimeMarker, 60000);

    // Fungsi list agenda per minggu
    const listModal = document.getElementById('listModal');
    let modalViewMode = 'weekly'; // 'weekly' atau 'daily'

    document.getElementById('btnListView').onclick = () => {
        listDateAnchor = new Date(selectedDate);
        document.getElementById('listModal').style.display = "block";
        renderListHTML();
    };

    document.getElementById('closeListModal').onclick = () => {
        listModal.style.display = "none";
    };

    document.getElementById('btnListWeekly').onclick = () => {
        modalViewMode = 'weekly';
        renderListHTML();
    };

    document.getElementById('btnListDaily').onclick = () => {
        modalViewMode = 'daily';
        renderListHTML();
    };

    function renderListHTML() {
        const titleElem = document.getElementById('listTitle');
        const rangeElem = document.getElementById('listRangeText');
        const containerElem = document.getElementById('listTableContainer');
        
        if (!titleElem || !rangeElem || !containerElem) return;
        
        const curr = new Date(listDateAnchor);
        const day = curr.getDay();
        let startDate, endDate, titleText;

        // 1. Logika Penentuan Range berdasarkan Mode
        if (modalViewMode === 'weekly') {
            const diff = curr.getDate() - day + (day === 0 ? -6 : 1);
            startDate = new Date(curr.setDate(diff));
            startDate.setHours(0,0,0,0);
            endDate = new Date(startDate);
            endDate.setDate(startDate.getDate() + 6);
            endDate.setHours(23,59,59);
            titleText = "List Agenda Mingguan";
        } else {
            startDate = new Date(curr.setHours(0,0,0,0));
            endDate = new Date(curr.setHours(23,59,59));
            titleText = "List Agenda Harian";
        }

        // Update UI Header
        const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        titleElem.innerText = titleText;
        rangeElem.innerText = modalViewMode === 'weekly'
            ? `${startDate.getDate()} ${monthNames[startDate.getMonth()]} — ${endDate.getDate()} ${monthNames[endDate.getMonth()]} ${endDate.getFullYear()}`
            : `${startDate.getDate()} ${monthNames[startDate.getMonth()]} ${startDate.getFullYear()}`;

        // 2. Filter & Sort
        const filteredList = allBookedData.filter(item => {
            if (item.tipePesanan === 'berulang') return false;
            let itemDate = item.tanggal.includes('/') ? new Date(item.tanggal.split('/').reverse().join('-')) : new Date(item.tanggal);
            return itemDate >= startDate && itemDate <= endDate;
        }).sort((a, b) => {
            const parseDate = (d) => d.includes('/') ? new Date(d.split('/').reverse().join('-')) : new Date(d);
            return parseDate(a.tanggal) - parseDate(b.tanggal) || a.jam.localeCompare(b.jam);
        });

        // 3. Bangun Tabel dengan Kolom "Jumlah Peserta"
        let tableHtml = `
            <table style="width: 100%; border-collapse: collapse; font-size: 0.85rem; table-layout: fixed;">
                <thead style="position: sticky; top: 0; background: #f1f5f9; z-index: 10;">
                    <tr>
                        <th style="padding: 10px; border-bottom: 2px solid #e2e8f0; width: 100px; text-align: center;">Hari/Tgl</th>
                        <th style="padding: 10px; border-bottom: 2px solid #e2e8f0; width: 85px; text-align: center;">Waktu</th>
                        <th style="padding: 10px; border-bottom: 2px solid #e2e8f0; text-align: left;">Agenda Acara</th>
                        <th style="padding: 10px; border-bottom: 2px solid #e2e8f0; width: 80px; text-align: center;">Ruang</th>
                        <th style="padding: 10px; border-bottom: 2px solid #e2e8f0; width: 70px; text-align: center;">Peserta</th>
                    </tr>
                </thead>
                <tbody>
                    ${filteredList.length > 0 ? filteredList.map(item => {
                        let dObj = item.tanggal.includes('/') ? new Date(item.tanggal.split('/').reverse().join('-')) : new Date(item.tanggal);
                        return `
                        <tr style="border-bottom: 1px solid #edf2f7;">
                            <td style="padding: 8px; text-align: center;">
                                <div style="font-weight: 700; color: #1a3a5f;">${dObj.toLocaleDateString('id-ID', { weekday: 'long' })}</div>
                                <div style="color: #64748b; font-size: 0.7rem;">${dObj.toLocaleDateString('id-ID', { day: '2-digit', month: '2-digit' })}</div>
                            </td>
                            <td style="padding: 8px; text-align: center;">
                                <span style="background: #f1f5f9; padding: 2px 5px; border-radius: 4px; font-weight: 600;">${item.jam}</span>
                            </td>
                            <td style="padding: 8px; text-align: left; vertical-align: top; white-space: normal; overflow: visible; word-wrap: break-word; word-break: break-word;">
                                <div style="font-weight: 600; color: #1e293b; line-height: 1.2;">${item.acara}</div>
                                <div style="font-size: 0.65rem; color: #94a3b8;">PIC: ${item.pic}</div>
                            </td>
                            <td style="padding: 8px; text-align: center;">
                                <span style="background: #e0f2fe; color: #0369a1; padding: 2px 6px; border-radius: 10px; font-weight: 700; font-size: 0.65rem;">${item.ruang}</span>
                            </td>
                            <td style="padding: 8px; text-align: center; color: #718096; font-style: italic;">
                                ${item.peserta || '-'}
                            </td>
                        </tr>`;
                    }).join('') : `<tr><td colspan="5" style="padding: 30px; text-align: center; color: #94a3b8;">Tidak ada agenda.</td></tr>`}
                </tbody>
            </table>`;

        document.getElementById('listTableContainer').innerHTML = tableHtml;

        // Update Button Active State
        document.getElementById('btnListWeekly').style.background = modalViewMode === 'weekly' ? '#2b6cb0' : '#edf2f7';
        document.getElementById('btnListWeekly').style.color = modalViewMode === 'weekly' ? 'white' : '#2d3748';
        document.getElementById('btnListDaily').style.background = modalViewMode === 'daily' ? '#2b6cb0' : '#edf2f7';
        document.getElementById('btnListDaily').style.color = modalViewMode === 'daily' ? 'white' : '#2d3748';

        // 4. Navigasi menyesuaikan mode
        document.getElementById('listPrev').onclick = (e) => {
            e.stopPropagation();
            modalViewMode === 'weekly' ? listDateAnchor.setDate(listDateAnchor.getDate() - 7) : listDateAnchor.setDate(listDateAnchor.getDate() - 1);
            renderListHTML();
        };
        document.getElementById('listNext').onclick = (e) => {
            e.stopPropagation();
            modalViewMode === 'weekly' ? listDateAnchor.setDate(listDateAnchor.getDate() + 7) : listDateAnchor.setDate(listDateAnchor.getDate() + 1);
            renderListHTML();
        };
    }

    function setupModalClose() {
        const closeListBtn = document.getElementById('closeListModal');
        if (closeListBtn) {
            closeListBtn.onclick = function() {
                listModal.style.display = "none";
            };
        }

        // Menutup jika klik di luar area putih modal
        window.addEventListener('click', function(event) {
            if (event.target == listModal) {
                listModal.style.display = "none";
            }
        });
    }
    setupModalClose();
});