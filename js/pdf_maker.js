const currentPage = window.location.pathname.split("/").pop();
document.querySelectorAll(".nav-links a").forEach((link) => {
  if (link.getAttribute("href") === currentPage) {
    link.classList.add("active");
  }
});

/* === Format Tanggal dd-mm-yyyy === */
function todayDDMMYYYY() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}-${mm}-${yyyy}`;
}

/* === Set Default Tanggal Gate Out (today) === */
window.onload = function () {
  go_tgl.value = todayDDMMYYYY();
  ss_tgl.value = todayDDMMYYYY();
  updateBox();
};

/* === AUTO COPY: Selesai Stuffing → Gate Out === */
ss_hari.addEventListener("input", () => {
  go_hari.value = ss_hari.value;
  updateBox();
});

ss_tgl.addEventListener("input", () => {
  go_tgl.value = ss_tgl.value;
  updateBox();
});

// === UPDATE MARKER REAL-TIME TANPA TOMBOL ===
[ss_hari, ss_tgl, ss_jam, go_hari, go_tgl, go_jam].forEach((input) => {
  input.addEventListener("input", updateBox);
});

let originalPdfBytes = null;
let currentViewport = null;

/* ------------ UPDATE MARKER PREVIEW ------------ */
function updateBox() {
  o_ss_hari.innerText = ss_hari.value;
  o_ss_tgl.innerText = ss_tgl.value;
  o_ss_jam.innerText = ss_jam.value;

  o_go_hari.innerText = go_hari.value;
  o_go_tgl.innerText = go_tgl.value;
  o_go_jam.innerText = go_jam.value;
}

/* ------------ LOAD PDF PREVIEW ------------ */
document.getElementById("pdf-file").addEventListener("change", function () {
  document.getElementById("marker-box").classList.remove("hidden");

  const file = this.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function () {
    originalPdfBytes = new Uint8Array(this.result);

    pdfjsLib.getDocument(originalPdfBytes).promise.then((pdf) => {
      pdf.getPage(1).then((page) => {
        currentViewport = page.getViewport({ scale: 1.3 });

        const canvas = document.getElementById("pdf-canvas");
        const ctx = canvas.getContext("2d");
        canvas.width = currentViewport.width;
        canvas.height = currentViewport.height;

        page.render({
          canvasContext: ctx,
          viewport: currentViewport,
        });
      });
    });
  };
  reader.readAsArrayBuffer(file);
});

async function downloadPDF() {
  if (!originalPdfBytes || !currentViewport) {
    alert("Upload PDF terlebih dahulu.");
    return;
  }

  const pdfDoc = await PDFLib.PDFDocument.load(originalPdfBytes);
  const page = pdfDoc.getPages()[0];
  const { width: pdfW, height: pdfH } = page.getSize();

  // pakai font standar Helvetica (mirip Arial)
  const helvFont = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);

  const canvas = document.getElementById("pdf-canvas");
  const marker = document.getElementById("marker-box");

  // posisi marker relatif terhadap canvas (supaya sama persis dengan HTML)
  const canvasRect = canvas.getBoundingClientRect();
  const markerRect = marker.getBoundingClientRect();

  const htmlLeft = markerRect.left - canvasRect.left;
  const htmlTop = markerRect.top - canvasRect.top;
  const htmlWidth = markerRect.width;
  const htmlHeight = markerRect.height;

  const s = currentViewport.scale;

  // konversi HTML px -> koordinat PDF (pt)
  const pdfX = htmlLeft / s;
  const pdfY = pdfH - htmlTop / s - htmlHeight / s;
  const boxW = htmlWidth / s;
  const boxH = htmlHeight / s;

  // ====== OUTLINE KOTAK ======
  page.drawRectangle({
    x: pdfX,
    y: pdfY,
    width: boxW,
    height: boxH,
    borderWidth: 2,
    borderColor: PDFLib.rgb(0, 0, 0),
  });

  // tinggi header (25px di CSS)
  const headerHeightPx = 25;
  const headerH = headerHeightPx / s;

  // garis header bawah
  page.drawLine({
    start: { x: pdfX, y: pdfY + boxH - headerH },
    end: { x: pdfX + boxW, y: pdfY + boxH - headerH },
    thickness: 2,
  });

  // garis tengah vertikal
  page.drawLine({
    start: { x: pdfX + boxW / 2, y: pdfY },
    end: { x: pdfX + boxW / 2, y: pdfY + boxH },
    thickness: 2,
  });

  // ====== HEADER TEXT (CENTER SEMPURNA) ======
  const headerSize = 11; // ≈ 14px
  const textLeft = "SELESAI STUFFING";
  const textRight = "GATE OUT";

  const leftWidth = helvFont.widthOfTextAtSize(textLeft, headerSize);
  const rightWidth = helvFont.widthOfTextAtSize(textRight, headerSize);

  const headerCenterY = pdfY + boxH - headerH / 2 - headerSize * 0.35; // posisinya di tengah header

  // kolom kiri
  page.drawText(textLeft, {
    x: pdfX + boxW / 4 - leftWidth / 2,
    y: headerCenterY,
    size: headerSize,
    font: helvFont,
  });

  // kolom kanan
  page.drawText(textRight, {
    x: pdfX + (3 * boxW) / 4 - rightWidth / 2,
    y: headerCenterY,
    size: headerSize,
    font: helvFont,
  });

  // ====== BODY TEXT (Hari / Tanggal / Jam) ======
  const bodySize = 10; // ≈ 13px
  const paddingTopPx = 22; // tambahkan padding supaya tidak nempel garis
  const rowGapPx = 23; // sedikit lebih renggang, mirip HTML

  // posisi baseline baris pertama
  const baseY = pdfY + boxH - headerH - paddingTopPx / s - bodySize * 0.2;

  const colPaddingPx = 10; // padding kiri
  const labelWidthPx = 55; // .marker-label { width:55px }
  const colonWidthPx = 10; // .marker-colon { width:10px }

  const colLeftX = pdfX + colPaddingPx / s;
  const colRightX = pdfX + boxW / 2 + colPaddingPx / s;

  function drawRow(colX, rowIndex, label, value) {
    const y = baseY - (rowIndex * rowGapPx) / s;

    const labelX = colX;
    const colonX = colX + labelWidthPx / s;
    const valueX = colX + (labelWidthPx + colonWidthPx) / s;

    // label
    page.drawText(label, {
      x: labelX,
      y,
      size: bodySize,
      font: helvFont,
    });

    // titik dua
    page.drawText(":", {
      x: colonX,
      y,
      size: bodySize,
      font: helvFont,
    });

    // value (boleh kosong)
    if (value) {
      page.drawText(value, {
        x: valueX,
        y,
        size: bodySize,
        font: helvFont,
      });
    }
  }

  // KIRI (SELESAI STUFFING — label saja, value kosong)
  drawRow(colLeftX, 0, "Hari", ss_hari.value || "");
  drawRow(colLeftX, 1, "Tanggal", ss_tgl.value || "");
  drawRow(colLeftX, 2, "Jam", ss_jam.value || "");

  // KANAN (GATE OUT — label + value dari input)
  drawRow(colRightX, 0, "Hari", go_hari.value || "");
  drawRow(colRightX, 1, "Tanggal", go_tgl.value || "");
  drawRow(colRightX, 2, "Jam", go_jam.value || "");

  // ====== SAVE PDF ======
  const finalPdf = await pdfDoc.save();
  const blob = new Blob([finalPdf], { type: "application/pdf" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "SPPB_Marked.pdf";
  link.click();
}
