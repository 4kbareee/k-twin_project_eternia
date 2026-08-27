// Slide Navigation and PPTX Export Logic
import PptxGenJS from 'pptxgenjs';

document.addEventListener("DOMContentLoaded", () => {
  const slides = document.querySelectorAll(".slide");
  const prevBtn = document.getElementById("prev-slide");
  const nextBtn = document.getElementById("next-slide");
  const currentIndexEl = document.getElementById("current-index");
  const totalSlidesEl = document.getElementById("total-slides");
  const progressBarFill = document.getElementById("progress-bar-fill");
  const fullscreenBtn = document.getElementById("toggle-fullscreen");
  const printBtn = document.getElementById("print-pdf");
  const downloadPptxBtn = document.getElementById("download-pptx");

  let currentSlide = 0;
  const totalSlides = slides.length;

  // Initialize display
  totalSlidesEl.textContent = totalSlides;
  updateUI();

  // Navigation functions
  function showSlide(index) {
    if (index >= 0 && index < totalSlides) {
      // Remove active class from current slide
      slides[currentSlide].classList.remove("active");
      
      // Update index
      currentSlide = index;
      
      // Add active class to new slide
      slides[currentSlide].classList.add("active");
      
      updateUI();
    }
  }

  function nextSlide() {
    if (currentSlide < totalSlides - 1) {
      showSlide(currentSlide + 1);
    }
  }

  function prevSlide() {
    if (currentSlide > 0) {
      showSlide(currentSlide - 1);
    }
  }

  function updateUI() {
    currentIndexEl.textContent = currentSlide + 1;
    
    // Progress Bar percentage
    const progressPercent = ((currentSlide + 1) / totalSlides) * 100;
    progressBarFill.style.width = `${progressPercent}%`;

    // Disable buttons at boundaries
    prevBtn.disabled = currentSlide === 0;
    nextBtn.disabled = currentSlide === totalSlides - 1;

    // Apply fading styling to disabled buttons
    prevBtn.style.opacity = currentSlide === 0 ? "0.3" : "1";
    prevBtn.style.cursor = currentSlide === 0 ? "not-allowed" : "pointer";
    nextBtn.style.opacity = currentSlide === totalSlides - 1 ? "0.3" : "1";
    nextBtn.style.cursor = currentSlide === totalSlides - 1 ? "not-allowed" : "pointer";
  }

  // Event Listeners for Nav buttons
  prevBtn.addEventListener("click", prevSlide);
  nextBtn.addEventListener("click", nextSlide);

  // Keyboard navigation
  document.addEventListener("keydown", (e) => {
    switch (e.key) {
      case "ArrowRight":
      case "Space":
      case " ":
        // Prevent default spacebar page scrolling
        if (e.key === " " || e.key === "Space") e.preventDefault();
        nextSlide();
        break;
      case "ArrowLeft":
        prevSlide();
        break;
      case "f":
      case "F":
        toggleFullscreen();
        break;
    }
  });

  // Swipe Gestures for Mobile Devices
  let touchStartX = 0;
  let touchEndX = 0;

  document.addEventListener("touchstart", (e) => {
    touchStartX = e.changedTouches[0].screenX;
  }, false);

  document.addEventListener("touchend", (e) => {
    touchEndX = e.changedTouches[0].screenX;
    handleSwipe();
  }, false);

  function handleSwipe() {
    const swipeThreshold = 50; // pixels
    if (touchEndX < touchStartX - swipeThreshold) {
      nextSlide(); // Swipe left -> Next
    }
    if (touchEndX > touchStartX + swipeThreshold) {
      prevSlide(); // Swipe right -> Prev
    }
  }

  // Fullscreen Management
  function toggleFullscreen() {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen()
        .then(() => {
          fullscreenBtn.querySelector("span").textContent = "Keluar Layar";
        })
        .catch(err => {
          console.error(`Gagal masuk mode layar penuh: ${err.message}`);
        });
    } else {
      document.exitFullscreen()
        .then(() => {
          fullscreenBtn.querySelector("span").textContent = "Layar Penuh";
        });
    }
  }

  fullscreenBtn.addEventListener("click", toggleFullscreen);

  // Print PDF Trigger
  printBtn.addEventListener("click", () => {
    window.print();
  });

  // PowerPoint Exporter Functionality
  function downloadPresentationAsPPTX() {
    // Show user some feedback
    const originalText = downloadPptxBtn.querySelector("span").textContent;
    downloadPptxBtn.querySelector("span").textContent = "Mengekspor...";
    downloadPptxBtn.disabled = true;

    try {
      let pres = new PptxGenJS();
      pres.layout = 'LAYOUT_16x9';

      const commonOptions = {
        fontFace: 'Segoe UI',
        color: '94A3B8'
      };

      const titleOptions = {
        fontFace: 'Segoe UI',
        fontSize: 26,
        bold: true,
        color: 'FFFFFF',
        x: 0.5,
        y: 0.4,
        w: 9.0,
        h: 0.6
      };

      const cardStyle = {
        fill: { color: '111129' },
        line: { color: '252542', width: 1 },
        margin: [10, 10, 10, 10]
      };

      // ----------------------------------------------------
      // SLIDE 1: Cover
      // ----------------------------------------------------
      let slide1 = pres.addSlide();
      slide1.background = { color: '0a0a16' };

      // Badge
      slide1.addText("SAAS PITCH DECK", {
        x: 0.5, y: 0.8, w: 9.0, h: 0.3,
        fontSize: 10, bold: true, color: 'C084FC', align: 'center', fontFace: 'Segoe UI'
      });

      // Main Title
      slide1.addText("K-Twin", {
        x: 0.5, y: 1.3, w: 9.0, h: 1.1,
        fontSize: 58, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Segoe UI'
      });

      // Subtitle
      slide1.addText("Your Organization's Knowledge, Personified.", {
        x: 0.5, y: 2.4, w: 9.0, h: 0.4,
        fontSize: 20, bold: true, color: 'A5B4FC', align: 'center', fontFace: 'Segoe UI'
      });

      // Tagline
      slide1.addText("Merevolusi transfer pengetahuan internal perusahaan dengan kecerdasan buatan terpersonalisasi.", {
        x: 1.0, y: 2.9, w: 8.0, h: 0.6,
        fontSize: 12, color: '94A3B8', align: 'center', fontFace: 'Segoe UI'
      });

      // Meta info boxes at bottom
      slide1.addText([
        { text: "CREATOR\n", options: { fontSize: 9, color: '94A3B8' } },
        { text: "Eternia Digital", options: { fontSize: 12, bold: true, color: 'FFFFFF' } }
      ], { x: 1.5, y: 4.1, w: 2.0, h: 0.8, align: 'center', fontFace: 'Segoe UI' });

      slide1.addText([
        { text: "TECH STACK\n", options: { fontSize: 9, color: '94A3B8' } },
        { text: "Gemini 3.5 API", options: { fontSize: 12, bold: true, color: 'FFFFFF' } }
      ], { x: 4.0, y: 4.1, w: 2.0, h: 0.8, align: 'center', fontFace: 'Segoe UI' });

      slide1.addText([
        { text: "TARGET MARKET\n", options: { fontSize: 9, color: '94A3B8' } },
        { text: "Enterprise & SMEs", options: { fontSize: 12, bold: true, color: 'FFFFFF' } }
      ], { x: 6.5, y: 4.1, w: 2.0, h: 0.8, align: 'center', fontFace: 'Segoe UI' });

      // ----------------------------------------------------
      // SLIDE 2: Masalah Utama Organisasi
      // ----------------------------------------------------
      let slide2 = pres.addSlide();
      slide2.background = { color: '0a0a16' };
      slide2.addText("Masalah Utama Organisasi", titleOptions);

      // Left Column
      slide2.addText("Seiring pertumbuhan perusahaan, mengelola dan mentransfer pengetahuan kerja (institutional knowledge) menjadi tantangan yang sangat mahal.", {
        x: 0.5, y: 1.2, w: 3.8, h: 1.6,
        fontSize: 13, color: '94A3B8', fontFace: 'Segoe UI'
      });

      // Left Stat Card
      slide2.addText([
        { text: "20%\n", options: { fontSize: 44, bold: true, color: 'FCA5A5' } },
        { text: "Waktu kerja staf habis hanya untuk mencari informasi internal.", options: { fontSize: 12, color: '94A3B8' } }
      ], {
        x: 0.5, y: 3.0, w: 3.8, h: 1.8,
        ...cardStyle, fontFace: 'Segoe UI'
      });

      // Right Column cards
      slide2.addText([
        { text: "Kehilangan Pengetahuan (Brain Drain)\n", options: { fontSize: 13, bold: true, color: 'FFFFFF' } },
        { text: "Ketika karyawan senior keluar (turnover), pengetahuan teknis mereka seringkali hilang bersama mereka.", options: { fontSize: 11, color: '94A3B8' } }
      ], {
        x: 4.7, y: 1.2, w: 4.8, h: 1.1,
        ...cardStyle, fontFace: 'Segoe UI'
      });

      slide2.addText([
        { text: "Dokumen Terkubur & Tidak Terbaca\n", options: { fontSize: 13, bold: true, color: 'FFFFFF' } },
        { text: "SOP, kebijakan HR, dan laporan keuangan ditumpuk di Google Drive tanpa ada yang membacanya secara efektif.", options: { fontSize: 11, color: '94A3B8' } }
      ], {
        x: 4.7, y: 2.5, w: 4.8, h: 1.1,
        ...cardStyle, fontFace: 'Segoe UI'
      });

      slide2.addText([
        { text: "Akses Data yang Siloed\n", options: { fontSize: 13, bold: true, color: 'FFFFFF' } },
        { text: "Staf teknis tidak tahu aturan cuti HR, dan staf baru kesulitan melacak riwayat perbaikan mesin tanpa senior.", options: { fontSize: 11, color: '94A3B8' } }
      ], {
        x: 4.7, y: 3.8, w: 4.8, h: 1.1,
        ...cardStyle, fontFace: 'Segoe UI'
      });

      // ----------------------------------------------------
      // SLIDE 3: Solusi K-Twin
      // ----------------------------------------------------
      let slide3 = pres.addSlide();
      slide3.background = { color: '0a0a16' };
      slide3.addText("Solusi: K-Twin (Knowledge Twin)", titleOptions);

      slide3.addText("K-Twin membuat 'Kembaran Digital Pengetahuan' yang cerdas untuk setiap divisi di perusahaan Anda.", {
        x: 0.5, y: 1.0, w: 9.0, h: 0.4,
        fontSize: 13, color: '94A3B8', align: 'center', fontFace: 'Segoe UI'
      });

      // Col 1
      slide3.addText([
        { text: "1. UNGGAH\n\n", options: { fontSize: 10, bold: true, color: 'C084FC' } },
        { text: "Unggah SOP & Data\n\n", options: { fontSize: 14, bold: true, color: 'FFFFFF' } },
        { text: "Cukup unggah dokumen kerja berformat PDF, Excel, atau CSV. Tidak perlu struktur database yang rumit.", options: { fontSize: 11, color: '94A3B8' } }
      ], {
        x: 0.5, y: 1.6, w: 2.8, h: 2.3,
        ...cardStyle, fontFace: 'Segoe UI'
      });

      // Col 2
      slide3.addText([
        { text: "2. PERSONALISASI\n\n", options: { fontSize: 10, bold: true, color: 'A5B4FC' } },
        { text: "Pilih Peran AI (Role)\n\n", options: { fontSize: 14, bold: true, color: 'FFFFFF' } },
        { text: "AI akan menyesuaikan gaya bahasanya: ramah (HR), teknis (Teknisi), atau akurat & formal (Finance).", options: { fontSize: 11, color: '94A3B8' } }
      ], {
        x: 3.6, y: 1.6, w: 2.8, h: 2.3,
        ...cardStyle, fontFace: 'Segoe UI'
      });

      // Col 3
      slide3.addText([
        { text: "3. INTERAKSI\n\n", options: { fontSize: 10, bold: true, color: 'F9A8D4' } },
        { text: "Tanya Jawab Instan\n\n", options: { fontSize: 14, bold: true, color: 'FFFFFF' } },
        { text: "Karyawan dapat bertanya apa saja dan mendapatkan referensi jawaban akurat langsung dari dokumen.", options: { fontSize: 11, color: '94A3B8' } }
      ], {
        x: 6.7, y: 1.6, w: 2.8, h: 2.3,
        ...cardStyle, fontFace: 'Segoe UI'
      });

      // Quote box
      slide3.addText("\"Seolah-olah memiliki asisten senior dari setiap divisi yang siap menjawab pertanyaan operasional Anda 24/7.\"", {
        x: 0.5, y: 4.1, w: 9.0, h: 0.8,
        fontSize: 12, italic: true, color: 'C7D2FE', align: 'center', valign: 'middle',
        fill: { color: '111129' }, line: { color: '252542', width: 1 }, fontFace: 'Segoe UI'
      });

      // ----------------------------------------------------
      // SLIDE 4: Fitur Utama
      // ----------------------------------------------------
      let slide4 = pres.addSlide();
      slide4.background = { color: '0a0a16' };
      slide4.addText("Fitur Utama SaaS K-Twin", titleOptions);

      // Left Visual (Chat Mockup)
      slide4.addText([
        { text: "🤖 K-Twin HR\n", options: { fontSize: 10, bold: true, color: 'C084FC' } },
        { text: "Halo! Saya K-Twin HR. Ada yang bisa saya bantu?\n\n", options: { fontSize: 10, color: '94A3B8' } },
        { text: "👤 User\n", options: { fontSize: 10, bold: true, color: 'FFFFFF' } },
        { text: "Berapa hari hak cuti melahirkan?\n\n", options: { fontSize: 10, color: 'A5B4FC' } },
        { text: "🤖 K-Twin HR\n", options: { fontSize: 10, bold: true, color: 'C084FC' } },
        { text: "Berdasarkan dokumen, Anda berhak atas 90 hari kalender cuti melahirkan.", options: { fontSize: 10, color: '94A3B8' } }
      ], {
        x: 0.5, y: 1.2, w: 4.2, h: 3.6,
        fill: { color: '0d0d21' }, line: { color: '252542', width: 1 },
        margin: [16, 16, 16, 16], fontFace: 'Segoe UI'
      });

      // Right Features
      slide4.addText([
        { text: "🛡️ Pemisahan Peran & Konteks Keamanan\n", options: { fontSize: 12, bold: true, color: 'FFFFFF' } },
        { text: "Data HR, Teknisi, dan Finance dipisah secara ketat untuk kepatuhan privasi data.", options: { fontSize: 10, color: '94A3B8' } }
      ], { x: 5.1, y: 1.2, w: 4.4, h: 0.8, fontFace: 'Segoe UI' });

      slide4.addText([
        { text: "📂 Parser Dokumen Multi-Format\n", options: { fontSize: 12, bold: true, color: 'FFFFFF' } },
        { text: "Mendukung parser file PDF, Excel (.xlsx, .xls), hingga CSV secara instan dan efisien.", options: { fontSize: 10, color: '94A3B8' } }
      ], { x: 5.1, y: 2.1, w: 4.4, h: 0.8, fontFace: 'Segoe UI' });

      slide4.addText([
        { text: "💬 Chat Streaming Responsif\n", options: { fontSize: 12, bold: true, color: 'FFFFFF' } },
        { text: "AI menjawab secara real-time (word-by-word) dengan parser Markdown otomatis.", options: { fontSize: 10, color: '94A3B8' } }
      ], { x: 5.1, y: 3.0, w: 4.4, h: 0.8, fontFace: 'Segoe UI' });

      slide4.addText([
        { text: "⚙️ Fleksibilitas BYOK (Bring Your Own Key)\n", options: { fontSize: 12, bold: true, color: 'FFFFFF' } },
        { text: "Bisa menggunakan API Key Google Gemini sendiri untuk kontrol biaya penuh dan privasi lokal.", options: { fontSize: 10, color: '94A3B8' } }
      ], { x: 5.1, y: 3.9, w: 4.4, h: 0.8, fontFace: 'Segoe UI' });

      // ----------------------------------------------------
      // SLIDE 5: Arsitektur & Alur Teknologi
      // ----------------------------------------------------
      let slide5 = pres.addSlide();
      slide5.background = { color: '0a0a16' };
      slide5.addText("Arsitektur & Alur Teknologi", titleOptions);

      slide5.addText("Bagaimana K-Twin memproses data menjadi asisten yang cerdas.", {
        x: 0.5, y: 1.0, w: 9.0, h: 0.4,
        fontSize: 13, color: '94A3B8', align: 'center', fontFace: 'Segoe UI'
      });

      // Flow 1
      slide5.addText([
        { text: "01\n\n", options: { fontSize: 18, bold: true, color: '6366F1' } },
        { text: "Input Dokumen\n\n", options: { fontSize: 13, bold: true, color: 'FFFFFF' } },
        { text: "PDF/Excel diunggah oleh user di browser.\n\n", options: { fontSize: 10, color: '94A3B8' } },
        { text: "FRONTEND PARSING", options: { fontSize: 8, bold: true, color: 'A5B4FC' } }
      ], {
        x: 0.5, y: 1.6, w: 2.0, h: 2.5,
        ...cardStyle, fontFace: 'Segoe UI'
      });

      // Flow 2
      slide5.addText([
        { text: "02\n\n", options: { fontSize: 18, bold: true, color: '6366F1' } },
        { text: "Ekstraksi Konteks\n\n", options: { fontSize: 13, bold: true, color: 'FFFFFF' } },
        { text: "Membaca teks PDF & konversi Excel ke CSV hemat token.\n\n", options: { fontSize: 10, color: '94A3B8' } },
        { text: "LOCAL PROCESSING", options: { fontSize: 8, bold: true, color: 'A5B4FC' } }
      ], {
        x: 2.85, y: 1.6, w: 2.0, h: 2.5,
        ...cardStyle, fontFace: 'Segoe UI'
      });

      // Flow 3
      slide5.addText([
        { text: "03\n\n", options: { fontSize: 18, bold: true, color: '6366F1' } },
        { text: "Prompt Dinamis\n\n", options: { fontSize: 13, bold: true, color: 'FFFFFF' } },
        { text: "Menggabungkan chat history, peran, konteks, & query.\n\n", options: { fontSize: 10, color: '94A3B8' } },
        { text: "PROMPT ENG", options: { fontSize: 8, bold: true, color: 'A5B4FC' } }
      ], {
        x: 5.2, y: 1.6, w: 2.0, h: 2.5,
        ...cardStyle, fontFace: 'Segoe UI'
      });

      // Flow 4
      slide5.addText([
        { text: "04\n\n", options: { fontSize: 18, bold: true, color: '8B5CF6' } },
        { text: "Generasi Respons\n\n", options: { fontSize: 13, bold: true, color: 'FFFFFF' } },
        { text: "Gemini 3.5 Flash menghasilkan respons streaming.\n\n", options: { fontSize: 10, color: '94A3B8' } },
        { text: "GOOGLE API STUDIO", options: { fontSize: 8, bold: true, color: 'C084FC' } }
      ], {
        x: 7.55, y: 1.6, w: 2.0, h: 2.5,
        fill: { color: '1A103C' }, line: { color: '6366F1', width: 1.5 },
        margin: [10, 10, 10, 10], fontFace: 'Segoe UI'
      });

      slide5.addText("Teknologi: Ultra-low Latency | Pemrosesan Client-Side (Privat) | Output Streaming", {
        x: 0.5, y: 4.4, w: 9.0, h: 0.4,
        fontSize: 11, color: '94A3B8', align: 'center', fontFace: 'Segoe UI'
      });

      // ----------------------------------------------------
      // SLIDE 6: Model Bisnis
      // ----------------------------------------------------
      let slide6 = pres.addSlide();
      slide6.background = { color: '0a0a16' };
      slide6.addText("Model Bisnis SaaS K-Twin", titleOptions);

      slide6.addText("Skema monetisasi yang dirancang fleksibel untuk efisiensi biaya infrastruktur.", {
        x: 0.5, y: 1.0, w: 9.0, h: 0.4,
        fontSize: 13, color: '94A3B8', align: 'center', fontFace: 'Segoe UI'
      });

      // Card 1
      slide6.addText([
        { text: "BYOK (Bring Your Own Key)\n\n", options: { fontSize: 13, bold: true, color: 'FFFFFF' } },
        { text: "Gratis / Tim Mandiri\n\n", options: { fontSize: 10, color: '94A3B8' } },
        { text: "✓ 3 Peran Dasar\n✓ Max 5 file per peran\n✓ Gunakan API Key sendiri\n✓ Penyimpanan lokal di browser", options: { fontSize: 10, color: 'FFFFFF' } }
      ], {
        x: 0.5, y: 1.6, w: 2.8, h: 3.3,
        ...cardStyle, fontFace: 'Segoe UI'
      });

      // Card 2 (Pro)
      slide6.addText([
        { text: "K-Twin Pro SaaS\n\n", options: { fontSize: 13, bold: true, color: 'FFFFFF' } },
        { text: "IDR 150K / user / bulan\n\n", options: { fontSize: 11, bold: true, color: 'C084FC' } },
        { text: "✓ Akses Kuota Terkelola\n✓ Max 10 Peran Kustom\n✓ Penyimpanan Cloud (50 file)\n✓ Kolaborasi antar anggota tim", options: { fontSize: 10, color: 'FFFFFF' } }
      ], {
        x: 3.6, y: 1.6, w: 2.8, h: 3.3,
        fill: { color: '1A103C' }, line: { color: '6366F1', width: 2 },
        margin: [12, 12, 12, 12], fontFace: 'Segoe UI'
      });

      // Card 3
      slide6.addText([
        { text: "K-Twin Enterprise\n\n", options: { fontSize: 13, bold: true, color: 'FFFFFF' } },
        { text: "Hubungi Kami\n\n", options: { fontSize: 10, color: '94A3B8' } },
        { text: "✓ Self-Hosted / On-Premise\n✓ Integrasi Database & ERP\n✓ SLA & Dukungan Teknis 24/7\n✓ Enkripsi E2E & Kepatuhan ISO", options: { fontSize: 10, color: 'FFFFFF' } }
      ], {
        x: 6.7, y: 1.6, w: 2.8, h: 3.3,
        ...cardStyle, fontFace: 'Segoe UI'
      });

      // ----------------------------------------------------
      // SLIDE 7: Proporsi Nilai & ROI
      // ----------------------------------------------------
      let slide7 = pres.addSlide();
      slide7.background = { color: '0a0a16' };
      slide7.addText("Proporsi Nilai & ROI Bisnis", titleOptions);

      // Left metrics
      slide7.addText([
        { text: "-85%\n", options: { fontSize: 24, bold: true, color: '22D3EE' } },
        { text: "Waktu Onboarding Karyawan Baru\n", options: { fontSize: 11, bold: true, color: 'FFFFFF' } },
        { text: "Karyawan baru tidak perlu menunggu mentor. Cukup bertanya ke Twin divisi.", options: { fontSize: 10, color: '94A3B8' } }
      ], {
        x: 0.5, y: 1.2, w: 4.2, h: 1.1,
        fill: { color: '111129' }, margin: [8, 10, 8, 10], fontFace: 'Segoe UI'
      });

      slide7.addText([
        { text: "10x\n", options: { fontSize: 24, bold: true, color: 'C084FC' } },
        { text: "Resolusi Masalah di Lapangan (Teknisi)\n", options: { fontSize: 11, bold: true, color: 'FFFFFF' } },
        { text: "Teknisi menemukan instruksi perbaikan dalam 3 detik, bukan 30 menit.", options: { fontSize: 10, color: '94A3B8' } }
      ], {
        x: 0.5, y: 2.5, w: 4.2, h: 1.1,
        fill: { color: '111129' }, margin: [8, 10, 8, 10], fontFace: 'Segoe UI'
      });

      slide7.addText([
        { text: "0%\n", options: { fontSize: 24, bold: true, color: '34D399' } },
        { text: "Risiko Kebocoran Dokumen Sensitif\n", options: { fontSize: 11, bold: true, color: 'FFFFFF' } },
        { text: "Penyimpanan lokal browser menjamin dokumen internal tidak keluar ke server luar.", options: { fontSize: 10, color: '94A3B8' } }
      ], {
        x: 0.5, y: 3.8, w: 4.2, h: 1.1,
        fill: { color: '111129' }, margin: [8, 10, 8, 10], fontFace: 'Segoe UI'
      });

      // Right box
      slide7.addText([
        { text: "Kenapa Memilih K-Twin?\n\n", options: { fontSize: 16, bold: true, color: 'FFFFFF' } },
        { text: "• Siap Pakai: ", options: { fontSize: 11, bold: true, color: 'C084FC' } },
        { text: "Tanpa fase deployment berminggu-minggu.\n\n", options: { fontSize: 11, color: '94A3B8' } },
        { text: "• Teknologi Terkini: ", options: { fontSize: 11, bold: true, color: 'C084FC' } },
        { text: "Memanfaatkan performa model Gemini 3.5 Flash.\n\n", options: { fontSize: 11, color: '94A3B8' } },
        { text: "• Hemat Token: ", options: { fontSize: 11, bold: true, color: 'C084FC' } },
        { text: "Parser cerdas memotong baris Excel, menghemat token API hingga 70%.\n\n", options: { fontSize: 11, color: '94A3B8' } },
        { text: "• Kustomisasi Peran: ", options: { fontSize: 11, bold: true, color: 'C084FC' } },
        { text: "Mendefinisikan peran AI kustom sesuai divisi kerja perusahaan.", options: { fontSize: 11, color: '94A3B8' } }
      ], {
        x: 5.1, y: 1.2, w: 4.4, h: 3.7,
        ...cardStyle, fontFace: 'Segoe UI'
      });

      // ----------------------------------------------------
      // SLIDE 8: Demo
      // ----------------------------------------------------
      let slide8 = pres.addSlide();
      slide8.background = { color: '0a0a16' };
      slide8.addText("Uji Coba Prototipe Sekarang", titleOptions);

      slide8.addText("Prototipe K-Twin sudah siap dijalankan untuk memvalidasi ide bisnis ini secara langsung.", {
        x: 0.5, y: 1.0, w: 9.0, h: 0.4,
        fontSize: 13, color: '94A3B8', align: 'center', fontFace: 'Segoe UI'
      });

      // Mockup
      slide8.addText([
        { text: "http://localhost:5173/\n\n", options: { fontSize: 11, color: '94A3B8' } },
        { text: "💬 Chatbot Utama | 📂 Riwayat Dokumen | ⚙️ Pengaturan API\n\n", options: { fontSize: 13, bold: true, color: 'FFFFFF' } },
        { text: "K-Twin by Eternia Digital\n", options: { fontSize: 12, bold: true, color: 'C084FC' } },
        { text: "Pilih Peran: HR / Teknisi / Finance. K-Twin siap membaca file Anda dan menjawab pertanyaan kerja secara instan.", options: { fontSize: 11, color: '94A3B8' } }
      ], {
        x: 1.5, y: 1.6, w: 7.0, h: 2.5,
        fill: { color: '0a0a14' }, line: { color: '252542', width: 1 },
        margin: [16, 20, 16, 20], fontFace: 'Segoe UI'
      });

      // CTA button
      slide8.addText("Buka Prototipe Chatbot (Tautan Aktif)", {
        x: 3.5, y: 4.4, w: 3.0, h: 0.5,
        fontSize: 11, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle',
        fill: { color: '6366F1' },
        hyperlink: { url: 'http://localhost:5173/' },
        fontFace: 'Segoe UI'
      });

      // ----------------------------------------------------
      // SLIDE 9: Roadmap
      // ----------------------------------------------------
      let slide9 = pres.addSlide();
      slide9.background = { color: '0a0a16' };
      slide9.addText("Visi & Peta Jalan (Roadmap)", titleOptions);

      slide9.addText("Langkah taktis Eternia Digital dalam mengembangkan K-Twin menjadi ekosistem pengetahuan.", {
        x: 0.5, y: 1.0, w: 9.0, h: 0.4,
        fontSize: 12, color: '94A3B8', align: 'center', fontFace: 'Segoe UI'
      });

      // Node 1
      slide9.addText([
        { text: "Q4 2026\n\n", options: { fontSize: 11, bold: true, color: 'C084FC' } },
        { text: "Fase 1: MVP\n\n", options: { fontSize: 12, bold: true, color: 'FFFFFF' } },
        { text: "Rilis prototipe web berbasis file lokal (PDF/Excel) dengan mode BYOK (versi saat ini).", options: { fontSize: 10, color: '94A3B8' } }
      ], { x: 0.5, y: 1.6, w: 2.0, h: 2.3, fill: { color: '111129' }, margin: 8, fontFace: 'Segoe UI' });

      // Node 2
      slide9.addText([
        { text: "Q1 2027\n\n", options: { fontSize: 11, bold: true, color: 'C084FC' } },
        { text: "Fase 2: Cloud RAG\n\n", options: { fontSize: 12, bold: true, color: 'FFFFFF' } },
        { text: "Integrasi database vektor cloud (Pinecone/Supabase) untuk membaca ribuan dokumen.", options: { fontSize: 10, color: '94A3B8' } }
      ], { x: 2.8, y: 1.6, w: 2.0, h: 2.3, fill: { color: '111129' }, margin: 8, fontFace: 'Segoe UI' });

      // Node 3 (Active)
      slide9.addText([
        { text: "Q2 2027\n\n", options: { fontSize: 11, bold: true, color: 'FFFFFF' } },
        { text: "Fase 3: Integrasi\n\n", options: { fontSize: 12, bold: true, color: 'FFFFFF' } },
        { text: "Integrasi bot resmi di Slack, Microsoft Teams, dan WhatsApp Business.", options: { fontSize: 10, color: '94A3B8' } }
      ], {
        x: 5.1, y: 1.6, w: 2.0, h: 2.3,
        fill: { color: '1A103C' }, line: { color: '6366F1', width: 1.5 },
        margin: 8, fontFace: 'Segoe UI'
      });

      // Node 4
      slide9.addText([
        { text: "Q3 2027\n\n", options: { fontSize: 11, bold: true, color: 'C084FC' } },
        { text: "Fase 4: Voice Twin\n\n", options: { fontSize: 12, bold: true, color: 'FFFFFF' } },
        { text: "AI text-to-speech kustom agar bisa diskusi dengan kembaran pengetahuan lewat suara.", options: { fontSize: 10, color: '94A3B8' } }
      ], { x: 7.4, y: 1.6, w: 2.0, h: 2.3, fill: { color: '111129' }, margin: 8, fontFace: 'Segoe UI' });

      // Footer
      slide9.addText("Eternia Digital — Empowering intelligence. Simplifying knowledge.", {
        x: 0.5, y: 4.4, w: 9.0, h: 0.5,
        fontSize: 12, italic: true, color: '94A3B8', align: 'center', valign: 'middle',
        fontFace: 'Segoe UI'
      });

      // ----------------------------------------------------
      // SAVE AND DOWNLOAD FILE
      // ----------------------------------------------------
      pres.writeFile({ fileName: 'K-Twin_SaaS_Pitch_Deck.pptx' })
        .then(() => {
          downloadPptxBtn.querySelector("span").textContent = originalText;
          downloadPptxBtn.disabled = false;
        })
        .catch(err => {
          console.error("Gagal mengunduh file PPTX:", err);
          downloadPptxBtn.querySelector("span").textContent = "Gagal!";
          setTimeout(() => {
            downloadPptxBtn.querySelector("span").textContent = originalText;
            downloadPptxBtn.disabled = false;
          }, 3000);
        });

    } catch (error) {
      console.error("Kesalahan pembuatan PPTX:", error);
      downloadPptxBtn.querySelector("span").textContent = originalText;
      downloadPptxBtn.disabled = false;
      alert("Terjadi kesalahan saat mengekspor presentasi.");
    }
  }

  // Hook Download button click
  if (downloadPptxBtn) {
    downloadPptxBtn.addEventListener("click", downloadPresentationAsPPTX);
  }
});
