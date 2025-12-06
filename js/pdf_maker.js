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
let extractedRegNumber = ""; // <-- FINAL NOMOR PENDAFTARAN

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
   LOAD PDF + EXTRACT NOMOR PENDAFTARAN (AKURASI 100%)
=========================================================== */
document.getElementById("pdf-file").addEventListener("change", function () {
  document.getElementById("marker-box").classList.remove("hidden");

  const file = this.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function () {
    originalPdfBytes = new Uint8Array(this.result);

    pdfjsLib.getDocument(originalPdfBytes).promise.then((pdf) => {
      pdf.getPage(1).then(async (page) => {
        currentViewport = page.getViewport({ scale: 1.3 });

        const canvas = document.getElementById("pdf-canvas");
        const ctx = canvas.getContext("2d");
        canvas.width = currentViewport.width;
        canvas.height = currentViewport.height;

        page.render({
          canvasContext: ctx,
          viewport: currentViewport,
        });

        /* === Extract text from PDF === */
        const textContent = await page.getTextContent();
        const items = textContent.items.map((i) => i.str.trim());

        extractedRegNumber = "";

        // 1. Cari label "Nomor Pendaftaran"
        let labelIdx = items.findIndex(
          (txt) => txt.replace(/\s+/g, "").toLowerCase() === "nomorpendaftaran"
        );

        // 2. Ambil angka sebelum label (INI NOMOR PENDAFTARAN YANG BENAR)
        if (labelIdx !== -1) {
          for (let i = labelIdx - 1; i >= 0; i--) {
            let part = items[i];
            if (!part) continue;

            if (/^\d+$/.test(part)) {
              extractedRegNumber = part;
              break;
            }
          }
        }

        // 3. Fallback jika gagal
        if (!extractedRegNumber) {
          const fallback = items.join(" ").match(/\b\d{5,20}\b/);
          extractedRegNumber = fallback ? fallback[0] : "UNKNOWN";
        }

        console.log("Nomor Pendaftaran FINAL:", extractedRegNumber);
      });
    });
  };

  reader.readAsArrayBuffer(file);
});

/* ===========================================================
   DOWNLOAD PDF DENGAN MARKER
=========================================================== */
async function downloadPDF() {
  if (!originalPdfBytes || !currentViewport) {
    alert("Upload PDF terlebih dahulu.");
    return;
  }

  const pdfDoc = await PDFLib.PDFDocument.load(originalPdfBytes);
  const page = pdfDoc.getPages()[0];
  const { width: pdfW, height: pdfH } = page.getSize();
  const helvFont = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);

  const canvas = document.getElementById("pdf-canvas");
  const marker = document.getElementById("marker-box");

  const canvasRect = canvas.getBoundingClientRect();
  const markerRect = marker.getBoundingClientRect();

  const htmlLeft = markerRect.left - canvasRect.left;
  const htmlTop = markerRect.top - canvasRect.top;
  const htmlWidth = markerRect.width;
  const htmlHeight = markerRect.height;

  const s = currentViewport.scale;

  const pdfX = htmlLeft / s;
  const pdfY = pdfH - htmlTop / s - htmlHeight / s;
  const boxW = htmlWidth / s;
  const boxH = htmlHeight / s;

  // Draw Box
  page.drawRectangle({
    x: pdfX,
    y: pdfY,
    width: boxW,
    height: boxH,
    borderWidth: 2,
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
  const headerSize = 11;
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
  const bodySize = 10;
  const paddingTopPx = 22;
  const rowGapPx = 23;

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

  /* === DOWNLOAD PDF === */
  const finalPdf = await pdfDoc.save();
  const blob = new Blob([finalPdf], { type: "application/pdf" });

  const safeName = extractedRegNumber.replace(/[\\/:*?"<>|]/g, "-") + ".pdf";

  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: `SPPB_${safeName}`,
      types: [
        {
          description: "PDF Document",
          accept: { "application/pdf": [".pdf"] },
        },
      ],
    });

    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();

    alert("File berhasil disimpan.");
  } catch (err) {
    console.log("Save canceled or failed:", err);
  }
}
