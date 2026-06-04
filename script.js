/* =========================================
   APLIKASI GENERATOR PRESENTASI POWERPOINT
   ========================================= */

const slidesContainer = document.getElementById("slides-container");
const addSlideButton = document.getElementById("add-slide-button");
const generatePresentationButton = document.getElementById("generate-presentation-button");
const educationLevelSelect = document.getElementById("education-level");

let slideCount = 1;

/* -----------------------------------------
   FUNGSI MEMBUAT SLIDE BARU DI WEB
----------------------------------------- */
function createNewSlide() {
    slideCount++;
    const newSlide = document.createElement("article");
    newSlide.classList.add("slide-card");
    newSlide.setAttribute("id", `slide-${slideCount}`);

    newSlide.innerHTML = `
        <h3 class="slide-title">Slide ${slideCount}</h3>
        <div class="form-group">
            <label for="slide-title-${slideCount}" class="input-label">Judul Materi</label>
            <input type="text" id="slide-title-${slideCount}" class="slide-title-input" name="slide_title_${slideCount}" placeholder="Masukkan judul materi">
        </div>
        <div class="form-group">
            <label for="slide-content-${slideCount}" class="input-label">Isi Materi</label>
            <textarea id="slide-content-${slideCount}" class="slide-content-textarea" name="slide_content_${slideCount}" rows="6" placeholder="Masukkan isi materi"></textarea>
        </div>
        <button type="button" class="delete-slide-button">Hapus Slide</button>
    `;

    slidesContainer.appendChild(newSlide);

    const deleteButton = newSlide.querySelector(".delete-slide-button");
    deleteButton.addEventListener("click", function () {
        newSlide.remove();
    });
}

addSlideButton.addEventListener("click", function () {
    createNewSlide();
});


/* -----------------------------------------
   FUNGSI MEMBUAT FILE POWERPOINT DENGAN DESAIN
----------------------------------------- */
function createPowerPoint(data) {
    let pres = new PptxGenJS();
    pres.layout = 'LAYOUT_16x9';

    // Desain Pekerjaan
    pres.defineSlideMaster({
        title: "MASTER_PEKERJAAN",
        background: { color: "F8FAFC" },
        objects: [
            { rect: { x: 0, y: 0, w: 0.15, h: "100%", fill: { color: "1E293B" } } },
            { rect: { x: 0.15, y: 0, w: 0.05, h: "100%", fill: { color: "D4AF37" } } },
            { text: { text: "Dokumen Presentasi Resmi", options: { x: 0.5, y: 5.3, w: 5, h: 0.3, fontSize: 10, color: "94A3B8" } } }
        ],
        slideNumber: { x: 9.5, y: 5.3, color: "94A3B8", fontSize: 10 }
    });

    // Desain Kuliah
    pres.defineSlideMaster({
        title: "MASTER_KULIAH",
        background: { color: "FFFFFF" },
        objects: [
            { rect: { x: 0, y: 0, w: "100%", h: 0.8, fill: { color: "0369A1" } } },
            { rect: { x: 0, y: 0.8, w: "100%", h: 0.05, fill: { color: "E2E8F0" } } },
            { text: { text: "Media Pembelajaran - Kelompok 6", options: { x: 0.5, y: 5.3, w: 5, h: 0.3, fontSize: 10, color: "64748B", bold: true } } }
        ],
        slideNumber: { x: 9.5, y: 5.3, color: "64748B", fontSize: 10 }
    });

    // Desain SMA
    pres.defineSlideMaster({
        title: "MASTER_SMA",
        background: { color: "FAFAFA" },
        objects: [
            { rect: { x: 7.5, y: 0, w: 2.5, h: 1, fill: { color: "DC2626" } } },
            { rect: { x: 7.3, y: 0, w: 0.15, h: 1, fill: { color: "111827" } } },
            { rect: { x: 0, y: 5.4, w: "100%", h: 0.2, fill: { color: "DC2626" } } }
        ],
        slideNumber: { x: 0.3, y: 5.15, color: "FFFFFF", fontSize: 11, bold: true }
    });

    // Desain SMP
    pres.defineSlideMaster({
        title: "MASTER_SMP",
        background: { color: "F0F9FF" },
        objects: [
            { rect: { x: 0, y: 0, w: 4, h: 0.5, fill: { color: "3B82F6" } } },
            { rect: { x: 4, y: 0, w: 6, h: 0.5, fill: { color: "FBBF24" } } },
            { rect: { x: 0, y: 0.5, w: 1, h: 0.1, fill: { color: "EF4444" } } }
        ],
        slideNumber: { x: 9.5, y: 5.2, color: "3B82F6", fontSize: 12, bold: true }
    });

    // Mengatur Tata Letak Teks Per Jenjang
    let masterToUse = "MASTER_KULIAH";
    let titleColor = "000000";
    let titleX = 0.5;
    let titleY = 0.5;
    let titleW = 9;

    if (data.educationLevel === "pekerjaan") {
        masterToUse = "MASTER_PEKERJAAN";
        titleColor = "1E293B"; 
        titleX = 0.6; 
    } else if (data.educationLevel === "kuliah") {
        masterToUse = "MASTER_KULIAH";
        titleColor = "FFFFFF"; 
        titleY = 0.15; 
    } else if (data.educationLevel === "sma") {
        masterToUse = "MASTER_SMA";
        titleColor = "111827"; 
        titleW = 7; 
    } else if (data.educationLevel === "smp") {
        masterToUse = "MASTER_SMP";
        titleColor = "1E3A8A"; 
        titleY = 0.7; 
    }

    // Memasukkan Teks ke Slide
    data.slides.forEach((slideData) => {
        let slide = pres.addSlide({ masterName: masterToUse });

        slide.addText(slideData.title, { 
            x: titleX, y: titleY, w: titleW, h: 0.8, 
            fontSize: 34, bold: true, color: titleColor, fontFace: "Arial" 
        });

        let contentY = (data.educationLevel === "kuliah") ? 1.2 : 1.5;

        slide.addText(slideData.content, { 
            x: 0.5, y: contentY, w: 9, h: 3.5, 
            fontSize: 22, color: "333333", align: "left", valign: "top", fontFace: "Arial"
        });
    });

    let fileName = `Presentasi_${data.educationLevel || 'Tanpa_Jenjang'}.pptx`;
    pres.writeFile({ fileName: fileName });
}


/* -----------------------------------------
   FUNGSI MENGAMBIL DATA DARI WEB
----------------------------------------- */
function collectPresentationData() {
    const allSlides = document.querySelectorAll(".slide-card");
    const slidesData = [];

    allSlides.forEach((slide, index) => {
        const titleInput = slide.querySelector(".slide-title-input");
        const contentTextarea = slide.querySelector(".slide-content-textarea");

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

    if (presentationData.slides.length > 0) {
        createPowerPoint(presentationData);
    } else {
        alert("Isi minimal satu judul dan materi sebelum membuat presentasi!");
    }
}


/* -----------------------------------------
   EVENT: KLIK TOMBOL BUAT PRESENTASI
----------------------------------------- */
generatePresentationButton.addEventListener("click", function () {
    collectPresentationData();
});

// Script untuk mengatur tampilan saat desain dipilih
document.getElementById('btn-template-elegan').onclick = function() { pilihTemplate('Monokrom Elegan'); };
document.getElementById('btn-template-biru').onclick = function() { pilihTemplate('Biru Akademik'); };
document.getElementById('btn-template-kreatif').onclick = function() { pilihTemplate('Kreatif Modern'); };

function pilihTemplate(nama) {
    // Tampilkan form pengisian materi
    document.getElementById('form-input-section').style.display = 'block';
    // Update teks nama template yang dipilih
    document.getElementById('nama-template-terpilih').innerText = nama;
    // Scroll layar otomatis ke form
    document.getElementById('form-input-section').scrollIntoView({ behavior: 'smooth' });
}
