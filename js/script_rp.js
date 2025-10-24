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
  const segel = getCell(wb, "HEADER", "CN2") || "";
  const kemUnit = getCell(wb, "KEMASAN", "C2");
  const kemQty = Number(getCell(wb, "KEMASAN", "D2")) || 0;
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
  return {
    aju,
    pengirim,
    bc,
    segel,
    kemasan: { unit: kemUnit, qty: kemQty },
    barang: { unit: barangUnit, total: barangTotal },
    tanggal: t ? new Date(t) : null,
    namaBarang, // simpan sebagai array
  };
}

// ---------- format tanggal dokumen ----------
function formatTanggalDokumen(arr) {
  if (!arr.length) return "";
  const sorted = [...new Set(arr.map((t) => t.getTime()))]
    .map((t) => new Date(t))
    .sort((a, b) => a - b);

  const sameMonth = sorted.every(
    (d) =>
      d.getMonth() === sorted[0].getMonth() &&
      d.getFullYear() === sorted[0].getFullYear()
  );

  if (sorted.length > 1 && sameMonth) {
    let consecutive = true;
    for (let i = 1; i < sorted.length; i++) {
      const diff = (sorted[i] - sorted[i - 1]) / (1000 * 3600 * 24);
      if (diff !== 1) {
        consecutive = false;
        break;
      }
    }
    if (consecutive) {
      const start = String(sorted[0].getDate()).padStart(2, "0");
      const end = String(sorted[sorted.length - 1].getDate()).padStart(2, "0");
      const mm = String(sorted[0].getMonth() + 1).padStart(2, "0");
      const yy = sorted[0].getFullYear();
      return `${start}–${end}/${mm}/${yy}`;
    }
  }
  return sorted.map(fmtDate).join(", ");
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
        <td>${d.aju}</td>
        <td>${d.pengirim}</td>
        <td>${d.bc || "-"}</td>
        <td>${d.kemasan.qty} ${d.kemasan.unit || ""}</td>
        <td>${d.barang.total} ${d.barang.unit || ""}</td>
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
  const pengirim = [
    ...new Set(dataArr.map((d) => d.pengirim).filter(Boolean)),
  ].join(" / ");
  const bcList = [...new Set(dataArr.map((d) => d.bc).filter(Boolean))].join(
    ", "
  );
  const segel = dataArr.find((d) => d.segel)?.segel || "";
  const kemasanMap = {};
  const barangMap = {};
  const tanggalArr = [];

  dataArr.forEach((d) => {
    if (d.kemasan.unit)
      kemasanMap[d.kemasan.unit] =
        (kemasanMap[d.kemasan.unit] || 0) + d.kemasan.qty;
    if (d.barang.unit)
      barangMap[d.barang.unit] =
        (barangMap[d.barang.unit] || 0) + d.barang.total;
    if (d.tanggal) tanggalArr.push(d.tanggal);
  });

  const kemasanText = Object.entries(kemasanMap)
    .map(([u, q]) => `${q} ${u}`)
    .join(" + ");
  const barangText = Object.entries(barangMap)
    .map(([u, q]) => `${q} ${u}`)
    .join(" + ");
  const tanggalDoc = formatTanggalDokumen(tanggalArr);
  const masukTxt = fmtDate(new Date($("masukTgl").value));

  return [
    "*BC 2.7 Masuk*",
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
  if (!selectedFiles.length) return alert("Pilih minimal 1 file Excel!");
  if (!$("jenisBarang").value)
    return alert("Pilih jenis barang terlebih dahulu!");
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
    alert("Terjadi kesalahan: " + err.message);
  } finally {
    $("processBtn").disabled = false;
    $("processBtn").textContent = "Proses File";
  }
});

$("copyBtn").addEventListener("click", () => {
  const text = $("result").value;
  if (!text) return alert("Belum ada hasil untuk disalin!");
  navigator.clipboard.writeText(text);
  alert("Teks berhasil disalin ke clipboard!");
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
