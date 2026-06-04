/* =========================================
   SISTEM GENERATOR PRESENTASI 15 HALAMAN
   ========================================= */

let templateTerpilih = "elegan"; // Default template

// =========================================
// 1. LOGIKA PEMILIHAN TEMPLATE
// =========================================
document.getElementById('btn-template-elegan').onclick = function() { pilihTemplate('elegan', 'Monokrom Elegan'); };
document.getElementById('btn-template-biru').onclick = function() { pilihTemplate('biru', 'Biru Akademik'); };
document.getElementById('btn-template-kreatif').onclick = function() { pilihTemplate('kreatif', 'Kreatif Modern'); };

function pilihTemplate(idTemplate, namaTemplate) {
    templateTerpilih = idTemplate;
    
    // Munculkan form 15 halaman
    document.getElementById('form-input-section').style.display = 'block';
    
    // Ubah label teks template yang dipilih
    const labelTemplate = document.getElementById('nama-template-terpilih');
    labelTemplate.innerText = namaTemplate;
    
    if(idTemplate === 'elegan') labelTemplate.style.color = '#111827';
    if(idTemplate === 'biru') labelTemplate.style.color = '#0369A1';
    if(idTemplate === 'kreatif') labelTemplate.style.color = '#F59E0B';

    // Otomatis scroll ke area form
    document.getElementById('form-input-section').scrollIntoView({ behavior: 'smooth' });
}

// =========================================
// 2. OTOMATIS MEMBUAT FORM HALAMAN 5 - 15
// =========================================
window.onload = function() {
    const container = document.getElementById("slides-container");

    // Halaman 5 sampai 7 (Materi KB 1 Lanjutan)
    for(let i = 5; i <= 7; i++) {
        container.appendChild(buatFormMateri(i, "Materi KB 1"));
    }

    // Halaman 8 (Judul KB 2)
    const kb2Card = document.createElement("article");
    kb2Card.className = "slide-card";
    kb2Card.style.borderLeft = "4px solid #F59E0B";
    kb2Card.innerHTML = `
        <h3 class="slide-title">Halaman 8: Pembuka Kegiatan Belajar 2</h3>
        <div class="form-group">
            <label class="input-label">Judul Kegiatan Belajar 2</label>
            <input type="text" id="kb2-title" class="slide-title-input" placeholder="Contoh: KB 2 - Strategi Pemasaran">
        </div>
    `;
    container.appendChild(kb2Card);

    // Halaman 9 sampai 15 (Materi KB 2)
    for(let i = 9; i <= 15; i++) {
        container.appendChild(buatFormMateri(i, "Materi KB 2"));
    }
};

// Fungsi cetak form materi berulang
function buatFormMateri(nomor, label) {
    const card = document.createElement("article");
    card.className = "slide-card";
    card.innerHTML = `
        <h3 class="slide-title">Halaman ${nomor}: ${label} (Lanjutan)</h3>
        <div class="form-group">
            <label class="input-label">Subjudul Materi</label>
            <input type="text" class="slide-title-input">
        </div>
        <div class="form-group">
            <label class="input-label">Isi Pembahasan</label>
            <textarea class="slide-content-textarea" rows="4"></textarea>
        </div>
        <div class="form-group">
            <label class="input-label">Tambahkan Gambar (Opsional)</label>
            <input type="file" accept="image/*" class="slide-image-input" style="border: none; padding: 0;">
        </div>
    `;
    return card;
}

// =========================================
// 3. FUNGSI MEMBACA GAMBAR JADI KODE (BASE64)
// =========================================
function getBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result);
        reader.onerror = error => reject(error);
    });
}

// =========================================
// 4. PROSES KLIK "BUAT PRESENTASI"
// =========================================
document.getElementById("generate-presentation-button").addEventListener("click", async function() {
    
    // Ubah teks tombol saat sedang loading memproses
    const originalText = this.innerText;
    this.innerText = "⏳ Sedang Merakit PPT...";
    this.disabled = true;

    try {
        let pres = new PptxGenJS();
        pres.layout = 'LAYOUT_16x9';

        // Tentukan warna tema berdasarkan pilihan desain
        let colorBgCover, colorAksen, colorBgContent;

        if (templateTerpilih === 'elegan') {
            colorBgCover = "111827";   // Abu-abu sangat gelap
            colorAksen = "9CA3AF";     // Perak elegan
            colorBgContent = "F9FAFB"; // Putih bersih
        } else if (templateTerpilih === 'biru') {
            colorBgCover = "0369A1";   // Biru akademik
            colorAksen = "FBBF24";     // Kuning emas
            colorBgContent = "F0F9FF"; // Biru sangat muda
        } else if (templateTerpilih === 'kreatif') {
            colorBgCover = "F59E0B";   // Oren/Kuning
            colorAksen = "FFFFFF";     // Putih
            colorBgContent = "FFFBEB"; // Kuning gading
        }

        // DESAIN: Master Cover & Judul KB
        pres.defineSlideMaster({
            title: "MASTER_COVER",
            background: { color: colorBgCover },
            objects: [
                { rect: { x: 0, y: 0, w: 0.2, h: "100%", fill: { color: colorAksen } } }
            ]
        });

        // DESAIN: Master Halaman Isi (Ada Header Gelap di Atas)
        pres.defineSlideMaster({
            title: "MASTER_CONTENT",
            background: { color: colorBgContent },
            objects: [
                { rect: { x: 0, y: 0, w: "100%", h: 1.1, fill: { color: colorBgCover } } },
                { rect: { x: 0, y: 1.1, w: "100%", h: 0.05, fill: { color: colorAksen } } }
            ],
            slideNumber: { x: 9.5, y: 5.2, color: "666666", fontSize: 10 }
        });

        // --------------------------------------------------
        // MEMBACA HTML DAN MENCETAK KE SLIDE SECARA BERURUTAN
        // --------------------------------------------------
        const allCards = document.querySelectorAll(".slide-card");
        
        // Ambil Data Cover
        const coverTitle = document.getElementById("cover-title").value.trim();
        const coverGroup = document.getElementById("cover-group").value.trim();
        const coverMembers = document.getElementById("cover-members").value.trim();

        // 1. Bikin Slide Cover
        let slideCover = pres.addSlide({ masterName: "MASTER_COVER" });
        slideCover.addText(coverTitle || "Judul Presentasi", { x: 0.5, y: 1.5, w: 9, h: 2, fontSize: 44, color: "FFFFFF", bold: true });
        slideCover.addText(coverGroup || "Nama Kelompok", { x: 0.5, y: 3.8, w: 9, h: 1, fontSize: 24, color: colorAksen });

        // 2. Bikin Slide Anggota (Bila Diisi)
        if(coverMembers !== "") {
            let slideAnggota = pres.addSlide({ masterName: "MASTER_CONTENT" });
            slideAnggota.addText("Daftar Anggota Kelompok", { x: 0.5, y: 0.2, w: 9, h: 0.8, fontSize: 30, color: "FFFFFF", bold: true });
            slideAnggota.addText(coverMembers, { x: 0.5, y: 1.5, w: 9, h: 3.5, fontSize: 22, color: "333333", valign: "top" });
        }

        // 3. Loop Semua Kartu Materi & KB
        for(let i = 0; i < allCards.length; i++) {
            let card = allCards[i];

            // Cek apakah ini Halaman Judul KB 1?
            let kb1TitleEl = card.querySelector("#kb1-title");
            if(kb1TitleEl && kb1TitleEl.value.trim() !== "") {
                let slide = pres.addSlide({ masterName: "MASTER_COVER" });
                slide.addText(kb1TitleEl.value.trim(), { x: 0.5, y: 2, w: 9, h: 1.5, fontSize: 40, color: "FFFFFF", bold: true, align: "center" });
                continue; // Lanjut ke halaman berikutnya
            }

            // Cek apakah ini Halaman Judul KB 2?
            let kb2TitleEl = card.querySelector("#kb2-title");
            if(kb2TitleEl && kb2TitleEl.value.trim() !== "") {
                let slide = pres.addSlide({ masterName: "MASTER_COVER" });
                slide.addText(kb2
