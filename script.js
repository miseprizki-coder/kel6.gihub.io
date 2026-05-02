/* =========================================
   APLIKASI GENERATOR PRESENTASI POWERPOINT
   ========================================= */

/* -----------------------------------------
   AMBIL ELEMEN YANG DIBUTUHKAN
----------------------------------------- */

const slidesContainer = document.getElementById("slides-container");
const addSlideButton = document.getElementById("add-slide-button");
const generatePresentationButton = document.getElementById("generate-presentation-button");
const educationLevelSelect = document.getElementById("education-level");

/* -----------------------------------------
   VARIABLE UNTUK MENGHITUNG JUMLAH SLIDE
   Slide pertama sudah ada dari HTML
----------------------------------------- */

let slideCount = 1;

/* -----------------------------------------
   FUNGSI MEMBUAT SLIDE BARU
----------------------------------------- */

function createNewSlide() {
    // Tambah nomor slide
    slideCount++;

    // Membuat elemen article untuk slide baru
    const newSlide = document.createElement("article");
    newSlide.classList.add("slide-card");
    newSlide.setAttribute("id", `slide-${slideCount}`);

    // Isi HTML untuk slide baru
    newSlide.innerHTML = `
        <h3 class="slide-title">Slide ${slideCount}</h3>

        <div class="form-group">
            <label for="slide-title-${slideCount}" class="input-label">
                Judul Materi
            </label>
            <input
                type="text"
                id="slide-title-${slideCount}"
                class="slide-title-input"
                name="slide_title_${slideCount}"
                placeholder="Masukkan judul materi"
            >
        </div>

        <div class="form-group">
            <label for="slide-content-${slideCount}" class="input-label">
                Isi Materi
            </label>
            <textarea
                id="slide-content-${slideCount}"
                class="slide-content-textarea"
                name="slide_content_${slideCount}"
                rows="6"
                placeholder="Masukkan isi materi"
            ></textarea>
        </div>

        <button
            type="button"
            class="delete-slide-button"
        >
            Hapus Slide
        </button>
    `;

    // Tambahkan slide baru ke container
    slidesContainer.appendChild(newSlide);

    // Ambil tombol hapus dari slide baru
    const deleteButton = newSlide.querySelector(".delete-slide-button");

    // Event hapus slide
    deleteButton.addEventListener("click", function () {
        newSlide.remove();
    });
}

/* -----------------------------------------
   EVENT: TOMBOL TAMBAH HALAMAN
----------------------------------------- */

addSlideButton.addEventListener("click", function () {
    createNewSlide();
});

/* -----------------------------------------
   FUNGSI MEMBUAT FILE POWERPOINT (PPTX) BENGAN DESAIN
----------------------------------------- */
function createPowerPoint(data) {
    let pres = new PptxGenJS();

    // ==========================================
    // 1. MEMBUAT DESAIN TEMPLATE (SLIDE MASTER)
    // ==========================================

    // Desain Pekerjaan (Formal, Elegan, Header Biru Dongker)
    pres.defineSlideMaster({
        title: "MASTER_PEKERJAAN",
        background: { color: "F8FAFC" },
        objects: [
            { rect: { x: 0, y: 0, w: "100%", h: 1.2, fill: { color: "1E293B" } } }, // Header gelap di atas
            { rect: { x: 0, y: 5.4, w: "100%", h: 0.2, fill: { color: "94A3B8" } } } // Garis abu-abu di bawah
        ]
    });

    // Desain Kuliah (Minimalis, Akademik, Garis Biru di Kiri)
    pres.defineSlideMaster({
        title: "MASTER_KULIAH",
        background: { color: "FFFFFF" },
        objects: [
            { rect: { x: 0, y: 0, w: 0.4, h: "100%", fill: { color: "0284C7" } } }, // Sidebar vertikal kiri
            { line: { x: 0.8, y: 1.2, w: 8.5, h: 0, line: { color: "0284C7", width: 2 } } } // Garis di bawah judul
        ]
    });

    // Desain SMA (Dinamis, Aksen Merah)
    pres.defineSlideMaster({
        title: "MASTER_SMA",
        background: { color: "FAFAFA" },
        objects: [
            { rect: { x: 0, y: 0, w: "100%", h: 0.8, fill: { color: "DC2626" } } }, // Header merah tipis
            { rect: { x: 0, y: 0.8, w: 0.2, h: "100%", fill: { color: "111827" } } } // Garis hitam tipis di kiri
        ]
    });

    // Desain SMP (Ceria, Aksen Kuning/Oranye)
    pres.defineSlideMaster({
        title: "MASTER_SMP",
        background: { color: "FFFBFC" },
        objects: [
            { rect: { x: 0, y: 0, w: "50%", h: 1.1, fill: { color: "F59E0B" } } }, // Blok warna kuning di atas kiri
            { rect: { x: "50%", y: 0, w: "50%", h: 1.1, fill: { color: "D97706" } } } // Blok warna oranye di atas kanan
        ]
    });

    // ==========================================
    // 2. MENYESUAIKAN TATA LETAK TEKS & TEMA
    // ==========================================
    
    let masterToUse = "MASTER_KULIAH"; // Default
    let titleColor = "0F172A"; // Default teks gelap
    let titleY = 0.4; // Posisi vertikal judul

    // Logika pemilihan desain
    if (data.educationLevel === "pekerjaan") {
        masterToUse = "MASTER_PEKERJAAN";
        titleColor = "FFFFFF"; // Teks putih karena menimpa header gelap
        titleY = 0.3; // Judul dinaikkan agar pas di dalam kotak header
    } else if (data.educationLevel === "kuliah") {
        masterToUse = "MASTER_KULIAH";
        titleColor = "0F172A";
        titleY = 0.5;
    } else if (data.educationLevel === "sma") {
        masterToUse = "MASTER_SMA";
        titleColor = "FFFFFF"; // Teks putih di dalam header merah
        titleY = 0.15;
    } else if (data.educationLevel === "smp") {
        masterToUse = "MASTER_SMP";
        titleColor = "FFFFFF"; 
        titleY = 0.3;
    }

    // ==========================================
    // 3. MEMASUKKAN TEKS KE DALAM SLIDE
    // ==========================================
    data.slides.forEach((slideData) => {
        // Buat slide dengan memanggil template desain di atas
        let slide = pres.addSlide({ masterName: masterToUse });

        // Tambahkan Teks Judul
        slide.addText(slideData.title, { 
            x: 0.8, y: titleY, w: 8.5, h: 0.8, 
            fontSize: 32, bold: true, color: titleColor, align: "left" 
        });

        // Tambahkan Teks Isi Materi (Posisi diatur agar tidak menabrak header)
        slide.addText(slideData.content, { 
            x: 0.8, y: 1.5, w: 8.5, h: 3.5, 
            fontSize: 20, color: "333333", align: "left", valign: "top" 
        });
    });

    // 4. Download PPTX
    let fileName = `Presentasi_${data.educationLevel || 'Tanpa_Jenjang'}.pptx`;
    pres.writeFile({ fileName: fileName });
}

/* -----------------------------------------
   UPDATE FUNGSI: MENGUMPULKAN DATA
----------------------------------------- */
// Modifikasi sedikit fungsi collectPresentationData yang lama 
// agar memanggil fungsi createPowerPoint di akhir
function collectPresentationData() {
    const allSlides = document.querySelectorAll(".slide-card");
    const slidesData = [];

    allSlides.forEach((slide, index) => {
        const titleInput = slide.querySelector(".slide-title-input");
        const contentTextarea = slide.querySelector(".slide-content-textarea");

        // Validasi: Jangan masukkan slide yang kosong
        if (titleInput.value.trim() !== "" || contentTextarea.value.trim() !== "") {
            slidesData.push({
                slideNumber: index + 1,
                title: titleInput.value.trim(),
                content: contentTextarea.value.trim()
            });
        }
    });

    const presentationData = {
        educationLevel: educationLevelSelect.value,
        slides: slidesData
    };

    // Panggil fungsi pembuat PowerPoint di sini!
    if (presentationData.slides.length > 0) {
        createPowerPoint(presentationData);
    } else {
        alert("Isi minimal satu judul dan materi sebelum membuat presentasi!");
    }
}

/* -----------------------------------------
   EVENT: TOMBOL BUAT PRESENTASI
----------------------------------------- */

generatePresentationButton.addEventListener("click", function () {
    collectPresentationData();
});