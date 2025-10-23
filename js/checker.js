// ---------- Helper Tabel Hasil ----------
function addResult(
  check,
  value,
  ref,
  isMatch,
  isQty = false,
  unitForRef = "",
  unitForData = undefined
) {
  if (unitForData === undefined) unitForData = unitForRef;
  const tbody = document.querySelector("#resultTable tbody");
  const row = document.createElement("tr");
  row.innerHTML = `
    <td>${check}</td>
    <td>${formatValue(value, isQty, unitForData)}</td>
    <td>${formatValue(ref, isQty, unitForRef)}</td>
    <td>${isMatch ? "Sama" : "Beda"}</td>
  `;
  row.classList.add(isMatch ? "match" : "mismatch");
  tbody.appendChild(row);
}

function applyFilter() {
  const filter = document.getElementById("filter").value;
  const rows = document.querySelectorAll("#resultTable tbody tr");
  rows.forEach((row) => {
    if (row.classList.contains("barang-header")) return;
    if (filter === "all") row.style.display = "";
    else if (filter === "sama")
      row.style.display = row.classList.contains("match") ? "" : "none";
    else if (filter === "beda")
      row.style.display = row.classList.contains("mismatch") ? "" : "none";
  });
}

// ---------- Fungsi Utama ----------
function checkAll(sheetPL, sheetINV, sheetsDATA, kurs, kontrakNo, kontrakTgl) {
  document.querySelector("#resultTable tbody").innerHTML = "";

  // Helper umum
  const normalize = (v) => {
    if (v === null || v === undefined) return "";
    if (!isNaN(v)) return parseFloat(v);
    return String(v).trim();
  };
  const isEqual = (a, b) => {
    const n1 = normalize(a),
      n2 = normalize(b);
    if (typeof n1 === "number" && typeof n2 === "number")
      return Math.abs(n1 - n2) < 0.01;
    return String(n1) === String(n2);
  };
  const isEqualStrict = (a, b) => (a || "") === (b || "");

  // ---------- Data PL ----------
  const { kemasanSum, bruttoSum, nettoSum, kemasanUnit } =
    hitungKemasanNWGW(sheetPL);

  // ---------- Data INV ----------
  const rangeINV = XLSX.utils.decode_range(sheetINV["!ref"]);
  const ptSelect = document.getElementById("ptSelect");
  const selectedPT = ptSelect.options[ptSelect.selectedIndex]?.text || "";

  const invCols = findHeaderColumns(sheetINV, {
    kode: selectedPT.includes("Shoetown")
      ? "STYLE"
      : selectedPT.includes("Long Rich")
      ? "STYLE.NO"
      : selectedPT.includes("Yih Quan")
      ? "SKU"
      : "MATERIAL CODE CUSTOMER",
    uraian: "ITEM NAME",
    qty: "QTY",
    cif: "AMOUNT",
    suratjalan: "SURAT JALAN",
  });

  const findInvoiceNo = (sheet) => {
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell = sheet[XLSX.utils.encode_cell({ r: R, c: C })];
        if (
          cell &&
          typeof cell.v === "string" &&
          cell.v.toUpperCase().includes("INVOICE NO")
        ) {
          const lines = cell.v.split(/\r?\n/);
          for (const line of lines) {
            if (line.toUpperCase().includes("INVOICE NO")) {
              const parts = line.split(":");
              if (parts.length > 1) return parts[1].trim();
            }
          }
        }
      }
    }
    return "";
  };

  // Hitung CIF
  let cifSum = 0;
  if (invCols.headerRow !== null && invCols.cif !== undefined) {
    for (let r = invCols.headerRow + 1; r <= rangeINV.e.r; r++) {
      const nomorSeri = getCellValue(sheetINV, "A" + (r + 1));
      if (!nomorSeri || isNaN(nomorSeri)) continue;
      cifSum +=
        parseFloat(
          getCellValue(sheetINV, XLSX.utils.encode_cell({ r, c: invCols.cif }))
        ) || 0;
    }
  }

  // Jenis Trx
  let jenisTransaksi = "";
  const n2Val = getCellValue(sheetsDATA.HEADER, "N2") || "";
  const selectedTrx = document.getElementById("jenisTrx")?.value?.trim() || "";

  switch (String(n2Val).trim()) {
    case "1":
      jenisTransaksi = "PENYERAHAN BKP";
      break;
    case "2":
      jenisTransaksi = "PENYERAHAN JKP";
      break;
    case "3":
      jenisTransaksi = "RETUR";
      break;
    case "4":
      jenisTransaksi = "NON PENYERAHAN";
      break;
    case "5":
      jenisTransaksi = "LAINNYA";
      break;
    default:
      jenisTransaksi = "TIDAK DIKETAHUI";
  }

  const isMatchTrx = jenisTransaksi.toUpperCase() === selectedTrx.toUpperCase();
  addResult("Jenis Transaksi", jenisTransaksi, selectedTrx, isMatchTrx);

  // Harga Penyerahan & Valuta
  const valuta = (getCellValue(sheetsDATA.HEADER, "CI2") || "").toUpperCase();
  const selectedValuta = (
    document.getElementById("valutaSelect")?.value || "USD"
  ).toUpperCase();
  const cifDraft = getCellValue(sheetsDATA.HEADER, "BU2");
  const cifMatch = isEqual(cifDraft, cifSum) && valuta === selectedValuta;

  addResult(
    "CIF",
    `${cifDraft} ${valuta}`,
    `${cifSum} ${selectedValuta}`,
    cifMatch,
    false
  );
  const hargaPenyerahan = getCellValue(sheetsDATA.HEADER, "BV2");
  addResult(
    "Harga Penyerahan",
    hargaPenyerahan,
    cifSum * kurs,
    isEqual(getCellValue(sheetsDATA.HEADER, "BV2"), cifSum * kurs)
  );

  // PPN 11%
  const dasarPengenaanPajak = getCellValue(sheetsDATA.HEADER, "BY2");
  addResult(
    "PPN 11%",
    dasarPengenaanPajak,
    cifSum * kurs * 0.11,
    Math.abs((dasarPengenaanPajak || 0) - cifSum * kurs * 0.11) < 0.01
  );

  // ---------- KEMASAN ----------
  const mapUnit = (u) => {
    if (!u) return "";
    const val = String(u).toUpperCase();
    if (val.includes("POLYBAG")) return "BG";
    if (val.includes("BOX")) return "BX";
    if (val.includes("CARTON")) return "CT";
    return val;
  };

  const kemasanUnitData = getCellValue(sheetsDATA.KEMASAN, "C2");
  const kemasanQtyData = getCellValue(sheetsDATA.KEMASAN, "D2");
  const kemasanUnitMapped = mapUnit(kemasanUnit);
  const kemasanUnitDataMapped = mapUnit(kemasanUnitData);
  const angkaMatch = isEqual(kemasanQtyData, kemasanSum);
  const unitMatch = kemasanUnitMapped === kemasanUnitDataMapped;

  addResult(
    "Total Kemasan",
    `${kemasanQtyData} ${kemasanUnitDataMapped}`,
    `${kemasanSum} ${kemasanUnitMapped}`,
    angkaMatch && unitMatch,
    true
  );

  // Total Brutto & Netto
  addResult(
    "Brutto",
    getCellValue(sheetsDATA.HEADER, "CB2"),
    bruttoSum,
    isEqual(getCellValue(sheetsDATA.HEADER, "CB2"), bruttoSum),
    false,
    "KG"
  );
  addResult(
    "Netto",
    getCellValue(sheetsDATA.HEADER, "CC2"),
    nettoSum,
    isEqual(getCellValue(sheetsDATA.HEADER, "CC2"), nettoSum),
    false,
    "KG"
  );

  // ---------- DOKUMEN ----------
  const invInvoiceNo = findInvoiceNo(sheetINV);
  const plInvoiceNo = findInvoiceNo(sheetPL);

  // 🔧 perubahan: fungsi ekstraksi diperbaiki agar bisa baca multiline & hapus teks tambahan
  // 🔧 ganti isi bagian dalam extractDateFromText()
  function extractDateFromText(text, label = "") {
    if (!text) {
      console.warn(`⚠️ [${label}] tidak ada teks DATE untuk diparse`);
      return "";
    }

    // 1️⃣ Normalisasi karakter & whitespace
    let src = String(text)
      .replace(/[\u00A0\u200B\uFEFF\u2003\u2002]/g, " ")
      .replace(/\r\n/g, "\n")
      .replace(/\t/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    console.log(
      `🟨 [${label}] teks setelah normalisasi:`,
      src.substring(0, 400)
    );

    // 2️⃣ Deteksi segmen yang mengandung kata DATE / Invoice Date / Packinglist Date
    //    Lebih longgar dan mendukung format: "DATE : 17 October 2025 DUE DATE :"
    const segPattern =
      /\b(?:Invoice\s*Date|Packing\s*List\s*Date|Packinglist\s*Date|DATE)\s*[:\-]?\s*([A-Za-z0-9\s,\/\-]+?)(?=(\bDUE\s*DATE\b|\bPO\s*NO\b|\bINVOICE\b|\bNo\s*Kontrak\b|\bTanggal\s*Kontrak\b|$))/i;

    const segMatch = src.match(segPattern);

    if (segMatch) {
      const rawCapture = segMatch[1].trim();
      console.log(`🔎 [${label}] segmen DATE cocok => '${rawCapture}'`);

      // Bersihkan bagian seperti "DUE DATE" bila tersisa
      const candidate = rawCapture
        .replace(/\bDUE\s*DATE\b.*$/i, "")
        .replace(/[^\w\s\-\/,\.]/g, " ")
        .replace(/\s+/g, " ")
        .trim();

      const parsed = tryParseDateCandidate(candidate, label);
      if (parsed) return parsed;

      console.warn(
        `⚠️ [${label}] segmen DATE ditemukan tapi gagal parse. candidate='${candidate}'`
      );
    } else {
      console.warn(
        `⚠️ [${label}] tidak menemukan segmen 'DATE' bertanda dalam teks`
      );
    }

    // 3️⃣ Fallback: cari tanggal umum di seluruh teks
    const globalDatePatterns = [
      /(\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\b)/, // 17 October 2025
      /([A-Za-z]+\s+\d{1,2},?\s+\d{4})/, // October 17, 2025
      /(\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}\b)/, // 17-10-2025
      /(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})/, // 2025-10-17
    ];

    for (const pat of globalDatePatterns) {
      const gm = src.match(pat);
      if (gm) {
        console.log(
          `🔁 [${label}] fallback menemukan tanggal global: '${gm[1]}'`
        );
        const parsed = tryParseDateCandidate(gm[1], label);
        if (parsed) return parsed;
      }
    }

    console.warn(`⚠️ [${label}] Gagal menemukan atau parse tanggal.`);
    return "";
  }

  function tryParseDateCandidate(raw, label = "") {
    if (!raw) return "";

    const r = raw
      .trim()
      .replace(/[,\u200B\uFEFF]/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    console.log(`   → [${label}] tryParseDateCandidate raw: '${r}'`);

    // dd Month yyyy
    let m = r.match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
    if (m) {
      const [_, d, mon, y] = m;
      const month = new Date(`${mon} 1, 2000`).getMonth() + 1;
      if (!isNaN(month)) {
        const iso = `${y}-${String(month).padStart(2, "0")}-${String(
          d
        ).padStart(2, "0")}`;
        console.log(`   ✅ parsed (text month) => ${iso}`);
        return iso;
      }
    }

    // Month dd yyyy
    m = r.match(/([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})/);
    if (m) {
      const [_, mon, d, y] = m;
      const month = new Date(`${mon} 1, 2000`).getMonth() + 1;
      if (!isNaN(month)) {
        const iso = `${y}-${String(month).padStart(2, "0")}-${String(
          d
        ).padStart(2, "0")}`;
        console.log(`   ✅ parsed (Month-first) => ${iso}`);
        return iso;
      }
    }

    // dd-mm-yyyy
    m = r.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) {
      const [_, dd, mm, yyyy] = m;
      const iso = `${yyyy}-${String(mm).padStart(2, "0")}-${String(dd).padStart(
        2,
        "0"
      )}`;
      console.log(`   ✅ parsed (numeric) => ${iso}`);
      return iso;
    }

    // yyyy-mm-dd
    m = r.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (m) {
      const [_, yyyy, mm, dd] = m;
      const iso = `${yyyy}-${String(mm).padStart(2, "0")}-${String(dd).padStart(
        2,
        "0"
      )}`;
      console.log(`   ✅ parsed (iso-ish) => ${iso}`);
      return iso;
    }

    // Fallback Date()
    const dObj = new Date(r);
    if (!isNaN(dObj)) {
      const yyyy = dObj.getFullYear();
      const mm = String(dObj.getMonth() + 1).padStart(2, "0");
      const dd = String(dObj.getDate()).padStart(2, "0");
      const iso = `${yyyy}-${mm}-${dd}`;
      console.log(`   ✅ parsed (Date() fallback) => ${iso}`);
      return iso;
    }

    return "";
  }

  // Ambil tanggal dari sheet DOKUMEN (Draft EXIM)
  function findDocDateByCode(sheet, code) {
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    for (let r = range.s.r; r <= range.e.r; r++) {
      const kode = getCellValueRC(sheet, r, 2); // kolom C = "KODE DOKUMEN"
      if (String(kode).trim() === String(code)) {
        return parseExcelDate(getCellValueRC(sheet, r, 4)); // kolom E = TANGGAL DOKUMEN
      }
    }
    return "";
  }

  const draftInvoiceDate = findDocDateByCode(sheetsDATA.DOKUMEN, "380");
  const draftPackinglistDate = findDocDateByCode(sheetsDATA.DOKUMEN, "217");

  // 🔧 perubahan: fungsi pencarian DATE di dalam file INV/PL kini membaca seluruh isi sel (multiline)
  function findDateText(sheet, label) {
    if (!sheet || !sheet["!ref"]) return "";
    const range = XLSX.utils.decode_range(sheet["!ref"]);

    console.log(
      `🔎 [${label}] mulai scan seluruh sheet (${range.e.r + 1} baris)`
    );

    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = sheet[addr];
        if (!cell || typeof cell.v !== "string") continue;

        const v = cell.v
          .replace(/[\u00A0\u200B\uFEFF\r\n\t]/g, " ")
          .replace(/\s+/g, " ")
          .trim();

        if (/DATE/i.test(v)) {
          console.log(`🟩 [${label}] ditemukan di ${addr}:`, v);
          return v;
        }
      }
    }

    // cek fallback: gabung semua teks
    const allText = Object.values(sheet)
      .filter((c) => c && typeof c.v === "string")
      .map((c) => c.v)
      .join(" ");
    const match = allText.match(/DATE\s*[:\-]?\s*([A-Za-z0-9 ,\/\-]+)/i);
    if (match) {
      console.log(`🔁 [${label}] fallback menemukan DATE: ${match[0]}`);
      return match[0];
    }

    console.warn(`⚠️ [${label}] tidak menemukan teks DATE di seluruh sheet`);
    return "";
  }

  const invDateText = findDateText(sheetINV, "Invoice");
  const plDateText = findDateText(sheetPL, "Packinglist");

  const invDateParsed = extractDateFromText(invDateText, "Invoice");
  const plDateParsed = extractDateFromText(plDateText, "Packinglist");

  addResult(
    "Invoice No.",
    getCellValue(sheetsDATA.DOKUMEN, "D2"),
    invInvoiceNo,
    isEqual(getCellValue(sheetsDATA.DOKUMEN, "D2"), invInvoiceNo)
  );
  addResult(
    "Invoice Date",
    draftInvoiceDate,
    invDateParsed,
    isEqual(draftInvoiceDate, invDateParsed)
  );
  addResult(
    "Packinglist No.",
    getCellValue(sheetsDATA.DOKUMEN, "D2"),
    plInvoiceNo,
    isEqual(getCellValue(sheetsDATA.DOKUMEN, "D2"), plInvoiceNo)
  );
  addResult(
    "Packinglist Date",
    draftPackinglistDate,
    plDateParsed,
    isEqual(draftPackinglistDate, plDateParsed)
  );

  let invSuratJalan = "";
  if (invCols.suratjalan !== undefined && invCols.headerRow !== null) {
    invSuratJalan = getCellValue(
      sheetINV,
      XLSX.utils.encode_cell({
        r: invCols.headerRow + 1,
        c: invCols.suratjalan,
      })
    );
  }

  // ---------- DELIVERY ORDER DATE ----------
  function findDocDateByCode(sheet, code) {
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    for (let r = range.s.r; r <= range.e.r; r++) {
      const kode = getCellValueRC(sheet, r, 2); // kolom C = "KODE DOKUMEN"
      if (String(kode).trim() === String(code)) {
        // kolom E = "TANGGAL DOKUMEN"
        return parseExcelDate(getCellValueRC(sheet, r, 4));
      }
    }
    return "";
  }

  const draftDeliveryOrderDate = findDocDateByCode(sheetsDATA.DOKUMEN, "640");

  addResult(
    "Delivery Order",
    getCellValue(sheetsDATA.DOKUMEN, "D5"),
    invSuratJalan,
    isEqual(getCellValue(sheetsDATA.DOKUMEN, "D5"), invSuratJalan)
  );

  addResult(
    "Delivery Order Date",
    draftDeliveryOrderDate,
    invDateParsed,
    isEqual(draftDeliveryOrderDate, invDateParsed)
  );

  // ---------- KONTRAK ----------
  // Fungsi universal untuk konversi tanggal Excel ke format yyyy-mm-dd
  // Fungsi universal untuk konversi nilai tanggal menjadi format yyyy-mm-dd
  function parseExcelDate(value) {
    if (!value) return "";

    // Jika berupa angka (serial date Excel)
    if (!isNaN(value)) {
      const serial = parseFloat(value);
      const utc_days = Math.floor(serial - 25569);
      const utc_value = utc_days * 86400;
      const date_info = new Date(utc_value * 1000);
      const year = date_info.getUTCFullYear();
      const month = String(date_info.getUTCMonth() + 1).padStart(2, "0");
      const day = String(date_info.getUTCDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    }

    // Jika string, bersihkan dulu
    let d = String(value).trim();

    // ---- Format yang sudah benar (yyyy-mm-dd) ----
    if (/^\d{4}-\d{2}-\d{2}$/.test(d)) return d;

    // ---- Format dd/mm/yyyy atau dd-mm-yyyy ----
    let match = d.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (match) {
      const [_, day, month, year] = match;
      return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
    }

    // ✅ ---- Format "1 October 2025" atau "01 October 2025" ----
    match = d.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
    if (match) {
      const [_, day, mon, year] = match;
      const monthIndex = new Date(`${mon} 1, 2000`).getMonth() + 1;
      if (!isNaN(monthIndex)) {
        return `${year}-${String(monthIndex).padStart(2, "0")}-${String(
          day
        ).padStart(2, "0")}`;
      }
    }

    // ✅ ---- Format "October 1, 2025" ----
    match = d.match(/^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})$/);
    if (match) {
      const [_, mon, day, year] = match;
      const monthIndex = new Date(`${mon} 1, 2000`).getMonth() + 1;
      if (!isNaN(monthIndex)) {
        return `${year}-${String(monthIndex).padStart(2, "0")}-${String(
          day
        ).padStart(2, "0")}`;
      }
    }

    // ---- Jika tidak cocok, kembalikan aslinya ----
    return d;
  }

  addResult(
    "Contract No.",
    getCellValue(sheetsDATA.DOKUMEN, "D4"),
    kontrakNo,
    isEqual(getCellValue(sheetsDATA.DOKUMEN, "D4"), kontrakNo)
  );

  const draftContractDateRaw = getCellValue(sheetsDATA.DOKUMEN, "E4");
  const draftContractDate = parseExcelDate(draftContractDateRaw);
  const kontrakTglFormatted = parseExcelDate(kontrakTgl);

  addResult(
    "Contract Date",
    draftContractDate,
    kontrakTglFormatted,
    isEqual(draftContractDate, kontrakTglFormatted)
  );
  const detectUnitFromPL = (sheet) => {
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    let foundPair = false;
    let foundPiece = false;

    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell = sheet[XLSX.utils.encode_cell({ r: R, c: C })];
        if (!cell || typeof cell.v !== "string") continue;
        const v = cell.v.trim().toUpperCase();
        if (
          v === "PAIRS" ||
          v === "PRS" ||
          v === "NPR" ||
          v.includes("PAIRS")
        ) {
          foundPair = true;
        }

        if (v === "PCS" || v === "PCE" || v === "PIECE") {
          foundPiece = true;
        }
      }
    }

    if (foundPair) return "NPR";
    if (foundPiece) return "PCE";
    return "NPR"; // Default
  };

  const detectedUnit = detectUnitFromPL(sheetPL);
  const rangeBarang = XLSX.utils.decode_range(sheetsDATA.BARANG["!ref"]);
  const plCols = findHeaderColumns(sheetPL, { nw: "NW", gw: "GW" });

  let barangCounter = 1;
  for (let r = 1; r <= rangeBarang.e.r; r++) {
    const kodeBarang = getCellValue(sheetsDATA.BARANG, "D" + (r + 1));
    if (!kodeBarang) continue;

    const rowINV = (invCols.headerRow || 0) + r;
    const rowPL = (plCols.headerRow || 0) + r;
    const tbody = document.querySelector("#resultTable tbody");
    const header = document.createElement("tr");
    header.classList.add("fw-bold", "barang-header");
    header.setAttribute("data-target", "barang-" + barangCounter);
    header.innerHTML = `<td colspan="4">BARANG KE ${barangCounter}</td>`;
    tbody.appendChild(header);

    const invKode = invCols.kode
      ? getCellValue(
          sheetINV,
          XLSX.utils.encode_cell({ r: rowINV, c: invCols.kode })
        )
      : "";
    addResult("Code", kodeBarang, invKode, isEqual(kodeBarang, invKode));

    const draftUraian = getCellValue(sheetsDATA.BARANG, "E" + (r + 1));
    const invUraian = invCols.uraian
      ? getCellValue(
          sheetINV,
          XLSX.utils.encode_cell({ r: rowINV, c: invCols.uraian })
        )
      : "";
    addResult(
      "Item Name",
      draftUraian,
      invUraian,
      isEqualStrict(draftUraian, invUraian)
    );

    // QTY Barang
    const draftQty = getCellValue(sheetsDATA.BARANG, "K" + (r + 1));
    const invQty = invCols.qty
      ? getCellValue(
          sheetINV,
          XLSX.utils.encode_cell({ r: rowINV, c: invCols.qty })
        )
      : "";
    const draftUnit = getCellValue(sheetsDATA.BARANG, "J2") || detectedUnit;
    const qtyMatch = isEqual(draftQty, invQty);
    const unitMatch = draftUnit === detectedUnit;
    addResult(
      "Quantity",
      draftQty,
      invQty,
      qtyMatch && unitMatch, // Sama hanya jika angka & unit sama
      true,
      detectedUnit, // Unit hasil deteksi dari PL
      draftUnit // Unit dari Draft EXIM
    );

    const draftNW = getCellValue(sheetsDATA.BARANG, "T" + (r + 1));
    const plNW = plCols.nw
      ? getCellValue(
          sheetPL,
          XLSX.utils.encode_cell({ r: rowPL, c: plCols.nw })
        )
      : "";
    addResult("NW", draftNW, plNW, isEqual(draftNW, plNW), false, "KG");

    const draftGW = getCellValue(sheetsDATA.BARANG, "U" + (r + 1));
    const plGW = plCols.gw
      ? getCellValue(
          sheetPL,
          XLSX.utils.encode_cell({ r: rowPL, c: plCols.gw })
        )
      : "";
    addResult("GW", draftGW, plGW, isEqual(draftGW, plGW), false, "KG");

    const draftCIF = getCellValue(sheetsDATA.BARANG, "Z" + (r + 1));
    const invCIF = invCols.cif
      ? getCellValue(
          sheetINV,
          XLSX.utils.encode_cell({ r: rowINV, c: invCols.cif })
        )
      : "";
    addResult(
      "Amount",
      `${draftCIF} ${valuta}`,
      `${invCIF} ${selectedValuta}`,
      isEqual(draftCIF, invCIF) && valuta === selectedValuta,
      false
    );

    barangCounter++;
  }

  // Collapsible rows
  document.querySelectorAll(".barang-header").forEach((header) => {
    header.addEventListener("click", () => {
      let next = header.nextElementSibling;
      while (next && !next.classList.contains("barang-header")) {
        next.style.display = next.style.display === "none" ? "" : "none";
        next = next.nextElementSibling;
      }
    });
  });
}
