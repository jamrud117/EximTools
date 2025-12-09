/* ===============================
   ACTIVE NAV LINK
================================ */
const currentPage = window.location.pathname.split("/").pop();
document.querySelectorAll(".nav-links a").forEach((link) => {
  if (link.getAttribute("href") === currentPage) {
    link.classList.add("active");
  }
});

/* ===============================
   FORMAT TANGGAL
================================ */
function todayDDMMYYYY() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}-${mm}-${yyyy}`;
}

/* ===============================
   GLOBAL VARIABLES
================================ */
let originalPdfBytes = null;
let currentViewport = null;
let extractedRegNumber = ""; // nomor pendaftaran untuk preview pertama
let pdfFiles = []; // semua file PDF yang diupload

/* ===============================
   DEFAULT VALUE
================================ */
window.onload = function () {
  go_tgl.value = todayDDMMYYYY();
  ss_tgl.value = todayDDMMYYYY();
  updateBox();
};

/* ===============================
   AUTO COPY SS → GO
================================ */
ss_hari.addEventListener("input", () => {
  go_hari.value = ss_hari.value;
  updateBox();
});

ss_tgl.addEventListener("input", () => {
  go_tgl.value = ss_tgl.value;
  updateBox();
});

// Update preview real-time
[ss_hari, ss_tgl, ss_jam, go_hari, go_tgl, go_jam].forEach((input) => {
  input.addEventListener("input", updateBox);
});

/* ===============================
   UPDATE LIVE PREVIEW BOX
================================ */
function updateBox() {
  o_ss_hari.innerText = ss_hari.value;
  o_ss_tgl.innerText = ss_tgl.value;
  o_ss_jam.innerText = ss_jam.value;

  o_go_hari.innerText = go_hari.value;
  o_go_tgl.innerText = go_tgl.value;
  o_go_jam.innerText = go_jam.value;
}

/* ===========================================================
   FUNGSI BANTU: EXTRACT NOMOR PENDAFTARAN DARI ITEMS TEKS
=========================================================== */
function extractRegFromItems(items) {
  let reg = "";

  // Normalisasi teks
  const trimmed = items.map((i) => (i || "").toString().trim());

  // 1. Cari label "Nomor Pendaftaran"
  let labelIdx = trimmed.findIndex(
    (txt) => txt.replace(/\s+/g, "").toLowerCase() === "nomorpendaftaran"
  );

  // 2. Ambil angka sebelum label (INI NOMOR PENDAFTARAN YANG BENAR)
  if (labelIdx !== -1) {
    for (let i = labelIdx - 1; i >= 0; i--) {
      let part = trimmed[i];
      if (!part) continue;

      if (/^\d+$/.test(part)) {
        reg = part;
        break;
      }
    }
  }

  // 3. Fallback jika gagal
  if (!reg) {
    const fallbackMatch = trimmed.join(" ").match(/\b\d{5,20}\b/);
    reg = fallbackMatch ? fallbackMatch[0] : "UNKNOWN";
  }

  return reg;
}

/* ===========================================================
   FUNGSI BANTU: EXTRACT NOMOR DARI BYTE PDF (UNTUK BATCH)
=========================================================== */
async function extractRegNumberFromBytes(uint8Array) {
  const pdf = await pdfjsLib.getDocument(uint8Array).promise;
  const page = await pdf.getPage(1);
  const textContent = await page.getTextContent();
  const items = textContent.items.map((i) => i.str || "");
  return extractRegFromItems(items);
}

/* ===========================================================
   LOAD PDF PERTAMA + PREVIEW + EXTRACT REG
=========================================================== */
document.getElementById("pdf-file").addEventListener("change", function () {
  const files = Array.from(this.files || []);
  pdfFiles = files;

  if (!pdfFiles.length) return;

  // Tampilkan marker-box
  document.getElementById("marker-box").classList.remove("hidden");

  // Load dan preview hanya PDF pertama
  loadFirstPreview(pdfFiles[0]);
});

function loadFirstPreview(file) {
  const reader = new FileReader();

  reader.onload = async function () {
    originalPdfBytes = new Uint8Array(this.result);

    const pdf = await pdfjsLib.getDocument(originalPdfBytes).promise;
    const page = await pdf.getPage(1);

    // Viewport untuk canvas
    currentViewport = page.getViewport({ scale: 1 });

    const canvas = document.getElementById("pdf-canvas");
    const ctx = canvas.getContext("2d");
    canvas.width = currentViewport.width;
    canvas.height = currentViewport.height;

    await page.render({
      canvasContext: ctx,
      viewport: currentViewport,
    });

    // Extract nomor pendaftaran untuk preview pertama (optional, untuk debugging)
    const textContent = await page.getTextContent();
    const items = textContent.items.map((i) => i.str || "");
    extractedRegNumber = extractRegFromItems(items);
    console.log("Nomor Pendaftaran PREVIEW:", extractedRegNumber);
  };

  reader.readAsArrayBuffer(file);
}

/* ===========================================================
   FUNGSI BANTU: HITUNG POSISI MARKER DI KOORDINAT PDF
   (PAKAI DOM MARKER + VIEWPORT PERTAMA)
=========================================================== */
function computeMarkerBoxForPage(page) {
  const { height: pdfH } = page.getSize();

  const canvas = document.getElementById("pdf-canvas");
  const marker = document.getElementById("marker-box");

  const canvasRect = canvas.getBoundingClientRect();
  const markerRect = marker.getBoundingClientRect();

  const htmlLeft = markerRect.left - canvasRect.left;
  const htmlTop = markerRect.top - canvasRect.top;
  const htmlWidth = markerRect.width;
  const htmlHeight = markerRect.height;

  // Skala diambil dari viewport pertama
  const s = currentViewport ? currentViewport.scale : 1;

  const pdfX = htmlLeft / s;
  const pdfY = pdfH - htmlTop / s - htmlHeight / s;
  const boxW = htmlWidth / s;
  const boxH = htmlHeight / s;

  return { pdfX, pdfY, boxW, boxH, s, pdfH };
}

/* ===========================================================
   FUNGSI: PROSES SATU PDF (TAMBAH MARKER)
=========================================================== */
async function processSinglePDF(uint8Array) {
  const pdfDoc = await PDFLib.PDFDocument.load(uint8Array);
  const page = pdfDoc.getPages()[0];
  const helvFont = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);

  // Hitung posisi marker di PDF
  const { pdfX, pdfY, boxW, boxH, s } = computeMarkerBoxForPage(page);

  // Draw Box
  page.drawRectangle({
    x: pdfX,
    y: pdfY,
    width: boxW,
    height: boxH,
    borderWidth: 1.2,
    borderColor: PDFLib.rgb(0, 0, 0),
  });

  // Header line
  const headerHeightPx = 25;
  const headerH = headerHeightPx / s;

  page.drawLine({
    start: { x: pdfX, y: pdfY + boxH - headerH },
    end: { x: pdfX + boxW, y: pdfY + boxH - headerH },
    thickness: 2,
  });

  page.drawLine({
    start: { x: pdfX + boxW / 2, y: pdfY },
    end: { x: pdfX + boxW / 2, y: pdfY + boxH },
    thickness: 2,
  });

  // Header text
  const headerSize = 12;
  const textLeft = "SELESAI STUFFING";
  const textRight = "GATE OUT";

  const leftW = helvFont.widthOfTextAtSize(textLeft, headerSize);
  const rightW = helvFont.widthOfTextAtSize(textRight, headerSize);

  const headerCenterY = pdfY + boxH - headerH / 2 - headerSize * 0.35;

  page.drawText(textLeft, {
    x: pdfX + boxW / 4 - leftW / 2,
    y: headerCenterY,
    size: headerSize,
    font: helvFont,
  });

  page.drawText(textRight, {
    x: pdfX + (3 * boxW) / 4 - rightW / 2,
    y: headerCenterY,
    size: headerSize,
    font: helvFont,
  });

  // Body text
  const bodySize = 11;
  const paddingTopPx = 18;
  const rowGapPx = 20;

  const baseY = pdfY + boxH - headerH - paddingTopPx / s - bodySize * 0.2;

  const colPaddingPx = 10;
  const labelWidthPx = 55;
  const colonWidthPx = 10;

  const colLeftX = pdfX + colPaddingPx / s;
  const colRightX = pdfX + boxW / 2 + colPaddingPx / s;

  function drawRow(colX, rowIndex, label, value) {
    const y = baseY - (rowIndex * rowGapPx) / s;

    const labelX = colX;
    const colonX = colX + labelWidthPx / s;
    const valueX = colX + (labelWidthPx + colonWidthPx) / s;

    page.drawText(label, { x: labelX, y, size: bodySize, font: helvFont });
    page.drawText(":", { x: colonX, y, size: bodySize, font: helvFont });

    if (value) {
      page.drawText(value, { x: valueX, y, size: bodySize, font: helvFont });
    }
  }

  // SS
  drawRow(colLeftX, 0, "Hari", ss_hari.value);
  drawRow(colLeftX, 1, "Tanggal", ss_tgl.value);
  drawRow(colLeftX, 2, "Jam", ss_jam.value);

  // GO
  drawRow(colRightX, 0, "Hari", go_hari.value);
  drawRow(colRightX, 1, "Tanggal", go_tgl.value);
  drawRow(colRightX, 2, "Jam", go_jam.value);

  const finalPdf = await pdfDoc.save();
  return finalPdf;
}

/* ===========================================================
   DOWNLOAD PDF: SINGLE = langsung, MULTI = ZIP
=========================================================== */
async function downloadAllPDF() {
  if (!pdfFiles.length) {
    alert("Upload beberapa file PDF terlebih dahulu.");
    return;
  }

  if (!currentViewport) {
    alert("Tunggu sampai preview PDF pertama tampil terlebih dahulu.");
    return;
  }

  try {
    /* =========================================
       CASE 1 — HANYA 1 FILE → DOWNLOAD LANGSUNG
    ========================================= */
    if (pdfFiles.length === 1) {
      const file = pdfFiles[0];
      const arrayBuffer = await file.arrayBuffer();
      const uint8Array = new Uint8Array(arrayBuffer);

      // extract nomor pendaftaran
      let regNumber = "";
      try {
        regNumber = await extractRegNumberFromBytes(uint8Array);
      } catch (err) {
        console.log("Gagal extract nomor pendaftaran:", err);
      }

      // proses PDF
      const processedBytes = await processSinglePDF(uint8Array);

      // nama file
      const baseName =
        regNumber && regNumber !== "UNKNOWN"
          ? regNumber
          : file.name.replace(/\.pdf$/i, "");

      const safeName = baseName.replace(/[\\/:*?"<>|]/g, "-") + ".pdf";

      // buat blob dan download
      const blob = new Blob([processedBytes], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.href = url;
      a.download = `SPPB_${safeName}`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      return; // STOP — tidak lanjut ke ZIP
    }

    /* =========================================
       CASE 2 — LEBIH DARI 1 FILE → ZIP
    ========================================= */
    const zip = new JSZip();

    for (const file of pdfFiles) {
      const arrayBuffer = await file.arrayBuffer();
      const uint8Array = new Uint8Array(arrayBuffer);

      // extract nomor pendaftaran per file
      let regNumber = "";
      try {
        regNumber = await extractRegNumberFromBytes(uint8Array);
      } catch (err) {
        console.log("Gagal extract nomor pendaftaran:", err);
      }

      const processedBytes = await processSinglePDF(uint8Array);

      const baseName =
        regNumber && regNumber !== "UNKNOWN"
          ? regNumber
          : file.name.replace(/\.pdf$/i, "");

      const safeName = baseName.replace(/[\\/:*?"<>|]/g, "-") + ".pdf";

      zip.file(`SPPB_${safeName}`, processedBytes);
    }

    const zipBlob = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(zipBlob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "SPPB_MARKED_ALL.zip";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  } catch (err) {
    console.error("Gagal membuat ZIP:", err);
    alert("Terjadi kesalahan saat membuat ZIP.");
  }
}
