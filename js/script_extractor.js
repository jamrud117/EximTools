const currentPage = window.location.pathname.split("/").pop();
document.querySelectorAll(".nav-links a").forEach((link) => {
  if (link.getAttribute("href") === currentPage) {
    link.classList.add("active");
  }
});

const barangCols = [
  "NO",
  "HS",
  "KODE BARANG",
  "SERI BARANG",
  "URAIAN",
  "KODE SATUAN",
  "JUMLAH SATUAN",
  "NETTO",
  "BRUTO",
  "CIF",
  "CIF RUPIAH",
  "NDPBM",
  "HARGA PENYERAHAN",
];

const ekstraksiCols = [
  "HS",
  "KODE BARANG",
  "URAIAN",
  "MEREK",
  "TIPE",
  "UKURAN",
  "SPESIFIKASI LAIN",
  "KODE SATUAN",
  "JUMLAH SATUAN",
  "KODE KEMASAN",
  "JUMLAH KEMASAN",
  "KODE DOKUMEN ASAL",
  "KODE KANTOR ASAL",
  "NOMOR DAFTAR ASAL",
  "TANGGAL DAFTAR ASAL",
  "NOMOR AJU ASAL",
  "SERI BARANG ASAL",
  "NETTO",
  "BRUTO",
  "VOLUME",
  "CIF",
  "CIF RUPIAH",
  "NDPBM",
  "HARGA PENYERAHAN",
];

let originalEkstrRows = [];
let currentEkstrRows = [];

const sheetToJSON = (s) =>
  XLSX.utils.sheet_to_json(s, { header: 1, raw: false });

const buildTable = (headers, rows) => {
  let html = "<table><thead><tr>";
  headers.forEach((h) => (html += `<th>${h}</th>`));
  html += "<th>Aksi</th></tr></thead><tbody>";
  rows.forEach((r, i) => {
    html += "<tr>";
    headers.forEach((_, j) => (html += `<td>${r[j] ?? ""}</td>`));
    html += `<td><button class="copyRowBtn" data-index="${i}">📋 Copy</button></td></tr>`;
  });
  html += "</tbody></table>";
  return html;
};

const attachCopyButtons = (id, data) => {
  document.querySelectorAll(`#${id} .copyRowBtn`).forEach((btn) =>
    btn.addEventListener("click", () => {
      navigator.clipboard.writeText(data[btn.dataset.index].join("\t"));
      btn.textContent = "✅ Copied!";
      setTimeout(() => (btn.textContent = "📋 Copy"), 1000);
    })
  );
};

const fadeUpdate = (el, html, after) => {
  el.classList.add("fade-out");
  setTimeout(() => {
    el.innerHTML = html;
    el.classList.remove("fade-out");
    el.classList.add("fade-in");
    if (after) after();
  }, 300);
};

function processWorkbook(wb) {
  const headerSheet = wb.Sheets["HEADER"];
  const barangSheet = wb.Sheets["BARANG"];
  const entitasSheet = wb.Sheets["ENTITAS"];
  if (!headerSheet || !barangSheet)
    return alert("Sheet HEADER atau BARANG tidak ditemukan!");

  // HEADER
  const header = {
    nomorAju: headerSheet["A2"]?.v || "",
    dokumen: headerSheet["B2"]?.v || "",
    kantor: headerSheet["C2"]?.v || "",
    daftar: headerSheet["CP2"]?.v || "",
    tanggal: headerSheet["CF2"]?.v || "",
  };

  let namaSupplier = "-";
  if (entitasSheet) {
    const entRaw = XLSX.utils.sheet_to_json(entitasSheet, { header: 1 });
    const headerRow = entRaw[0].map((h) =>
      (h || "").toString().trim().toUpperCase()
    );
    const kodeIdx = headerRow.indexOf("KODE ENTITAS");
    const namaIdx = headerRow.indexOf("NAMA ENTITAS");
    if (kodeIdx >= 0 && namaIdx >= 0) {
      const supplierRow = entRaw.find((row, i) => i > 0 && row[kodeIdx] == 3);
      if (supplierRow) namaSupplier = supplierRow[namaIdx] || "-";
    }
  }

  const headerLabels = {
    nomorAju: "Nomor Aju Asal",
    dokumen: "Kode Dokumen Asal",
    kantor: "Kode Kantor Asal",
    daftar: "Nomor Daftar Asal",
    tanggal: "Tanggal Daftar Asal",
  };

  document.getElementById("headerContent").innerHTML = `
          <table>
            <tr><th>Informasi</th><th>Data</th></tr>
            <tr><td>Nama Supplier</td><td>${namaSupplier}</td></tr>
            ${Object.entries(header)
              .map(
                ([k, v]) =>
                  `<tr><td>${headerLabels[k] || k}</td><td>${v}</td></tr>`
              )
              .join("")}
          </table>
        `;

  // DATA BARANG
  const raw = sheetToJSON(barangSheet);
  const headers = raw[0];
  const data = raw.slice(1);
  const idx = (n) =>
    headers.findIndex((h) => (h || "").toUpperCase().trim() === n);

  const barangRows = data.map((r, i) =>
    barangCols.map((c) => {
      if (c === "NO") return i + 1; // ← otomatis isi nomor urut
      return idx(c) >= 0 ? r[idx(c)] ?? "" : "";
    })
  );

  document.getElementById("barangCard").style.display = "block";
  document.getElementById("barangTableWrap").innerHTML = buildTable(
    barangCols,
    barangRows
  );
  attachCopyButtons("barangTableWrap", barangRows);

  // DATA EKSTRAKSI
  // DATA EKSTRAKSI (dengan logika CIF=0 → Harga Penyerahan, NDPBM default 1, CIF Rupiah dari Excel)
  const ekstrRows = data.map((r) => {
    const cifExcel = parseFloat(r[idx("CIF")]) || 0;
    const ndpbmExcel = parseFloat(r[idx("NDPBM")]) || 0;
    const hargaExcel = parseFloat(r[idx("HARGA PENYERAHAN")]) || 0;
    const cifRupiahExcel = parseFloat(r[idx("CIF RUPIAH")]) || 0;

    // Default NDPBM = 1 jika 0
    const ndpbm = ndpbmExcel === 0 ? 1 : ndpbmExcel;

    // Jika CIF Excel = 0 → pakai Harga Penyerahan dari Excel
    const cifFinal = cifExcel === 0 ? hargaExcel : cifExcel;

    // Jika CIF Excel = 0 → Harga Penyerahan juga = harga Excel
    const hargaFinal = cifExcel === 0 ? hargaExcel : cifExcel * ndpbm;

    return ekstraksiCols.map((c) => {
      if (c === "KODE DOKUMEN ASAL") return header.dokumen;
      if (c === "KODE KANTOR ASAL") return header.kantor;
      if (c === "NOMOR DAFTAR ASAL") return header.daftar;
      if (c === "TANGGAL DAFTAR ASAL") return header.tanggal;
      if (c === "NOMOR AJU ASAL") return header.nomorAju;
      if (c === "SERI BARANG ASAL") return r[idx("SERI BARANG")] ?? "";

      if (c === "CIF") return formatNumber(cifFinal);
      if (c === "HARGA PENYERAHAN") return formatNumber(hargaFinal);

      // Ambil CIF Rupiah dari Excel, jangan 0
      if (c === "CIF RUPIAH") return formatNumber(cifRupiahExcel);

      if (c === "NDPBM") return formatNumber(ndpbm);

      const i = idx(c);
      return i >= 0 ? r[i] ?? "" : "";
    });
  });

  originalEkstrRows = JSON.parse(JSON.stringify(ekstrRows));
  currentEkstrRows = JSON.parse(JSON.stringify(ekstrRows));

  const wrap = document.getElementById("ekstraksiTableWrap");
  document.getElementById("ekstraksiCard").style.display = "block";
  fadeUpdate(wrap, buildTable(ekstraksiCols, ekstrRows), () =>
    attachCopyButtons("ekstraksiTableWrap", ekstrRows)
  );

  // Dropdown Barang
  const select = document.getElementById("barangSelect");
  select.innerHTML = "";
  select.appendChild(new Option("TAMPILKAN SEMUA", "all"));
  ekstrRows.forEach((_, i) =>
    select.appendChild(new Option(`BARANG KE ${i + 1}`, i))
  );

  select.addEventListener("change", () => {
    const v = select.value;
    if (v === "all")
      fadeUpdate(wrap, buildTable(ekstraksiCols, currentEkstrRows), () =>
        attachCopyButtons("ekstraksiTableWrap", currentEkstrRows)
      );
    else
      fadeUpdate(wrap, buildTable(ekstraksiCols, [currentEkstrRows[v]]), () =>
        attachCopyButtons("ekstraksiTableWrap", [currentEkstrRows[v]])
      );
  });
}
function formatNumber(value) {
  return Number.isInteger(value) ? value.toString() : value.toFixed(2);
}

// Terapkan Quantity
function applyQuantity() {
  const qty = parseFloat(document.getElementById("quantityInput").value);
  const select = document.getElementById("barangSelect");
  const index = parseInt(select.value);

  if (select.value === "all") return alert("Pilih barang tertentu!");
  if (isNaN(qty)) return alert("Masukkan quantity valid!");

  const row = [...currentEkstrRows[index]];
  const qtyAwal = parseFloat(row[8]) || 1;
  const cifAwal = parseFloat(row[20]) || 0;
  const cifRupiahAwal = parseFloat(row[21]) || 0;
  const ndpbm = parseFloat(row[22]) || 1;

  const unitCIF = cifAwal / qtyAwal;
  const unitCIFRupiah = cifRupiahAwal / qtyAwal;

  const cifBaru = unitCIF * qty;
  const cifRupiahBaru = unitCIFRupiah * qty;
  const hargaPenyerahanBaru = cifBaru * ndpbm;

  row[8] = formatNumber(qty);
  row[20] = formatNumber(cifBaru);
  row[21] = formatNumber(cifRupiahBaru);
  row[23] = formatNumber(hargaPenyerahanBaru);

  currentEkstrRows[index] = row;

  fadeUpdate(
    document.getElementById("ekstraksiTableWrap"),
    buildTable(ekstraksiCols, [row]),
    () => attachCopyButtons("ekstraksiTableWrap", [row])
  );
}

function resetData() {
  currentEkstrRows = JSON.parse(JSON.stringify(originalEkstrRows));
  document.getElementById("quantityInput").value = "";
  document.getElementById("barangSelect").value = "all";
  fadeUpdate(
    document.getElementById("ekstraksiTableWrap"),
    buildTable(ekstraksiCols, currentEkstrRows),
    () => attachCopyButtons("ekstraksiTableWrap", currentEkstrRows)
  );
}

document
  .getElementById("applyQuantityBtn")
  .addEventListener("click", applyQuantity);
document.getElementById("resetBtn").addEventListener("click", resetData);

document.getElementById("fileInput").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => {
    const wb = XLSX.read(new Uint8Array(ev.target.result), {
      type: "array",
    });
    processWorkbook(wb);
  };
  reader.readAsArrayBuffer(file);
});
