const currentPage = window.location.pathname.split("/").pop();
document.querySelectorAll(".nav-links a").forEach((link) => {
  if (link.getAttribute("href") === currentPage) {
    link.classList.add("active");
  }
});

// ---------- utilitas umum ----------
const $ = (id) => document.getElementById(id);
const fmtDate = (d) =>
  `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(
    2,
    "0"
  )}/${d.getFullYear()}`;

// ---------- pembacaan file ----------
async function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        resolve({ file, wb });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}
const fmtNum = (n) =>
  typeof n === "number"
    ? n.toLocaleString("id-ID")
    : Number(n || 0).toLocaleString("id-ID");

// ---------- ekstraksi data ----------
function getCell(wb, sheet, addr) {
  const s = wb.Sheets[sheet];
  return s && s[addr] ? s[addr].v : undefined;
}

function getEntitas(wb) {
  const s = wb.Sheets["ENTITAS"];
  if (!s) return "";
  const data = XLSX.utils.sheet_to_json(s, { header: 1 });
  if (!data.length) return "";
  let kodeIdx = -1,
    namaIdx = -1;
  for (let i = 0; i < data[0].length; i++) {
    const val = String(data[0][i] || "")
      .trim()
      .toUpperCase();
    if (val === "KODE ENTITAS") kodeIdx = i;
    if (val === "NAMA ENTITAS") namaIdx = i;
  }
  for (let r = 1; r < data.length; r++) {
    const kode = data[r][kodeIdx];
    if (kode === 3 || String(kode).trim() === "3")
      return String(data[r][namaIdx] || "").trim();
  }
  return "";
}

function extractDataFromWorkbook(file, wb) {
  const pengirim = getEntitas(wb);
  const bc = getCell(wb, "HEADER", "CP2") || "";
  const segel = getCell(wb, "HEADER", "BC2") || "";
  // ---- ambil semua data kemasan ----
  const sheetKemasan = wb.Sheets["KEMASAN"];
  let kemasanMap = {};

  if (sheetKemasan && sheetKemasan["!ref"]) {
    const dataKemasan = XLSX.utils.sheet_to_json(sheetKemasan, { header: 1 });
    const header = dataKemasan[0] || [];
    let kodeIdx = -1,
      jumlahIdx = -1;

    // cari kolom berdasarkan nama header
    for (let i = 0; i < header.length; i++) {
      const val = String(header[i] || "")
        .trim()
        .toUpperCase();
      if (val === "KODE KEMASAN") kodeIdx = i;
      if (val === "JUMLAH KEMASAN") jumlahIdx = i;
    }

    // fallback jika header tidak ditemukan (anggap C=kode, D=jumlah)
    if (kodeIdx === -1) kodeIdx = 2;
    if (jumlahIdx === -1) jumlahIdx = 3;

    // iterasi semua baris data (mulai dari baris ke-2)
    for (let r = 1; r < dataKemasan.length; r++) {
      const kode = String(dataKemasan[r][kodeIdx] || "").trim();
      const qty = Number(dataKemasan[r][jumlahIdx]) || 0;
      if (!kode) continue;
      kemasanMap[kode] = (kemasanMap[kode] || 0) + qty;
    }
  }

  // jika ada beberapa jenis kemasan, ambil semuanya
  const kemUnit = Object.keys(kemasanMap).length
    ? Object.keys(kemasanMap)[0]
    : "";
  const kemQty = kemasanMap[kemUnit] || 0;

  const barangUnit = getCell(wb, "BARANG", "J2");
  let barangTotal = 0;
  let namaBarang = [];

  const sheetBarang = wb.Sheets["BARANG"];
  if (sheetBarang && sheetBarang["!ref"]) {
    const dataBarang = XLSX.utils.sheet_to_json(sheetBarang, {
      header: 1,
    });

    // cari kolom header "URAIAN"
    const headerRow = dataBarang[0] || [];
    const uraianIdx = headerRow.findIndex(
      (h) => String(h).trim().toUpperCase() === "URAIAN"
    );

    // jika kolom ditemukan, ambil semua nilainya
    if (uraianIdx !== -1) {
      const namaArr = [];
      for (let r = 1; r < dataBarang.length; r++) {
        const nama = dataBarang[r][uraianIdx];
        if (nama) namaArr.push(String(nama).trim());
      }
      // hilangkan duplikat
      namaBarang = [...new Set(namaArr)];
    }

    // hitung total barang (kolom J)
    const range = XLSX.utils.decode_range(sheetBarang["!ref"]);
    for (let R = range.s.r; R <= range.e.r; ++R) {
      const addr = XLSX.utils.encode_cell({ c: 10, r: R });
      const cell = sheetBarang[addr];
      if (cell && !isNaN(cell.v)) barangTotal += Number(cell.v);
    }
  }

  const t = getCell(wb, "HEADER", "CF2");
  const aju = getCell(wb, "HEADER", "A2") || "";
  const jenistrx = getCell(wb, "HEADER", "N2");
  let jenisTransaksi = "";
  const n2Val = jenistrx;

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

  return {
    jenistrx: jenisTransaksi,
    aju,
    pengirim,
    bc,
    segel,
    kemasan: kemasanMap,
    barang: { unit: barangUnit, total: barangTotal },
    tanggal: t ? new Date(t) : null,
    namaBarang,
  };
}

// ---------- format tanggal dokumen (versi cerdas) ----------
function formatTanggalDokumen(arr) {
  if (!arr.length) return "";

  // Hilangkan duplikat & urutkan
  const sorted = [...new Set(arr.map((t) => t.getTime()))]
    .map((t) => new Date(t))
    .sort((a, b) => a - b);

  const groups = [];
  let start = sorted[0];
  let end = sorted[0];

  for (let i = 1; i < sorted.length; i++) {
    const prev = end;
    const current = sorted[i];
    const diff = (current - prev) / (1000 * 3600 * 24);

    if (
      diff === 1 && // tanggal berurutan
      current.getMonth() === start.getMonth() &&
      current.getFullYear() === start.getFullYear()
    ) {
      end = current; // masih satu rentang
    } else {
      groups.push([start, end]);
      start = current;
      end = current;
    }
  }
  groups.push([start, end]);

  // Format tiap grup
  const formattedGroups = groups.map(([s, e]) => {
    const dd1 = String(s.getDate()).padStart(2, "0");
    const dd2 = String(e.getDate()).padStart(2, "0");
    const mm = String(s.getMonth() + 1).padStart(2, "0");
    const yy = s.getFullYear();

    if (s.getTime() === e.getTime()) {
      // tanggal tunggal
      return `${dd1}/${mm}/${yy}`;
    } else {
      // rentang tanggal berurutan
      return `${dd1}-${dd2}/${mm}/${yy}`;
    }
  });

  return formattedGroups.join(", ");
}

// ---------- tampilan UI ----------
function updateFileList(files) {
  $("fileList").textContent =
    files.length > 0
      ? files.map((f) => f.name).join(", ")
      : "Belum ada file dipilih.";
}

function renderPreview(dataArr) {
  const tbody = $("previewTable").querySelector("tbody");
  tbody.innerHTML = dataArr
    .map(
      (d) => `
      <tr>
        <td>${d.jenistrx}</td>
        <td>${d.aju}</td>
        <td>${d.pengirim}</td>
        <td>${d.bc || "-"}</td>
        <td>${Object.entries(d.kemasan)
          .map(([unit, qty]) => `${fmtNum(qty)} ${unit}`)
          .join("<br>")}</td>
        <td>${fmtNum(d.barang.total)} ${d.barang.unit || ""}</td>
        <td>${d.tanggal ? fmtDate(d.tanggal) : ""}</td>
        <td>${
          d.namaBarang && d.namaBarang.length ? d.namaBarang.join("<br>") : "-"
        }</td>
      </tr>`
    )
    .join("");
  $("tableWrap").style.display = "block";
}

function generateResultText(dataArr) {
  const jenisTrx = [
    ...new Set(dataArr.map((d) => d.jenistrx).filter(Boolean)),
  ].join(" | ");
  const pengirim = [
    ...new Set(dataArr.map((d) => d.pengirim).filter(Boolean)),
  ].join(" | ");
  const bcList = [...new Set(dataArr.map((d) => d.bc).filter(Boolean))].join(
    ", "
  );
  const segel = dataArr.find((d) => d.segel)?.segel || "";
  const tanggalArr = [];

  // Akumulasi kemasan dan barang
  const kemasanMap = {};
  const barangMap = {};

  dataArr.forEach((d) => {
    // kemasan
    if (d.kemasan && typeof d.kemasan === "object") {
      for (const [unit, qty] of Object.entries(d.kemasan)) {
        kemasanMap[unit] = (kemasanMap[unit] || 0) + qty;
      }
    }

    // barang
    if (d.barang.unit) {
      barangMap[d.barang.unit] =
        (barangMap[d.barang.unit] || 0) + d.barang.total;
    }

    // tanggal dokumen
    if (d.tanggal) tanggalArr.push(d.tanggal);
  });

  const kemasanText = Object.entries(kemasanMap)
    .map(([u, q]) => `${fmtNum(q)} ${u}`)
    .join(" + ");

  const barangText = Object.entries(barangMap)
    .map(([u, q]) => `${fmtNum(q)} ${u}`)
    .join(" + ");

  const tanggalDoc = formatTanggalDokumen(tanggalArr);
  const masukTxt = fmtDate(new Date($("masukTgl").value));

  return [
    "*BC 2.7 Masuk*",
    `Jenis Transaksi : ${jenisTrx}`,
    `Pengirim : ${pengirim}`,
    `No BC 2.7 : ${bcList}`,
    `No Segel : ${segel}`,
    `Jumlah kemasan : ${kemasanText}`,
    `Jenis Barang : ${$("jenisBarang").value}`,
    `Jumlah barang : ${barangText}`,
    `Tanggal Dokumen : ${tanggalDoc}`,
    `Masuk Tgl : ${masukTxt}`,
  ].join("\n");
}

// ---------- event handler ----------
$("masukTgl").value = new Date().toISOString().slice(0, 10);
let selectedFiles = [];

$("files").addEventListener("change", (e) => {
  selectedFiles = Array.from(e.target.files);
  updateFileList(selectedFiles);
});

$("processBtn").addEventListener("click", async () => {
  if (!selectedFiles.length)
    return Swal.fire({
      icon: "error",
      title: "Oops...",
      scrollbarPadding: false,
      text: "Pilih minimal 1 file Excel!",
    });
  if (!$("jenisBarang").value)
    return Swal.fire({
      icon: "error",
      title: "Oops...",
      scrollbarPadding: false,
      text: "Pilih jenis barang terlebih dahulu!",
    });
  $("processBtn").disabled = true;
  $("processBtn").textContent = "Memproses...";

  try {
    const workbooks = await Promise.all(selectedFiles.map(readWorkbook));
    const extracted = workbooks.map(({ file, wb }) =>
      extractDataFromWorkbook(file, wb)
    );
    renderPreview(extracted);
    $("result").value = generateResultText(extracted);
  } catch (err) {
    console.error(err);
    Swal.fire({
      icon: "error",
      title: "Oops...",
      text: "Terjadi kesalahan: " + err.message,
      scrollbarPadding: false,
    });
  } finally {
    $("processBtn").disabled = false;
    $("processBtn").textContent = "Proses File";
  }
});

$("copyBtn").addEventListener("click", () => {
  const text = $("result").value;
  if (!text)
    return Swal.fire({
      icon: "error",
      title: "Oops...",
      text: "Belum ada hasil untuk disalin!",
      scrollbarPadding: false,
    });
  navigator.clipboard.writeText(text);
  Swal.fire({
    position: "top-mid",
    icon: "success",
    title: "Teks berhasil disalin ke clipboard!",
    showConfirmButton: false,
    scrollbarPadding: false,
    timer: 1500,
  });
});

$("clearBtn").addEventListener("click", () => {
  $("files").value = "";
  selectedFiles = [];
  updateFileList([]);
  $("previewTable").querySelector("tbody").innerHTML = "";
  $("tableWrap").style.display = "none";
  $("result").value = "";
  $("jenisBarang").value = "";
  $("masukTgl").value = new Date().toISOString().slice(0, 10);
});
