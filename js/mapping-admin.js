let mappings = {};

function load() {
  const saved = localStorage.getItem("customerMappings");
  if (saved) mappings = JSON.parse(saved);
  render();
}

function saveMapping() {
  const pt = document.getElementById("mapPT").value.trim();
  const code = document.getElementById("mapKode").value.trim();
  const uraian = document.getElementById("mapUraian").value.trim();
  const qty = document.getElementById("mapQty").value.trim();
  const cif = document.getElementById("mapCIF").value.trim();
  const suratjalan = document.getElementById("mapSJ").value.trim();

  // VALIDASI GENERIK — semua field wajib diisi
  const fields = [
    { value: pt, label: "Nama customer" },
    { value: code, label: "Code customer" },
    { value: uraian, label: "Nama item customer" },
    { value: qty, label: "Header quantity customer" },
    { value: cif, label: "Header CIF customer" },
    { value: suratjalan, label: "Header surat jalan customer" },
  ];

  const emptyField = fields.find((f) => !f.value || f.value.trim() === "");

  if (emptyField) {
    alert(`${emptyField.label} wajib diisi!`);
    return; // stop proses
  }

  // Jika semua terisi → simpan mapping
  mappings[pt] = {
    kode: code,
    uraian: uraian,
    qty: qty,
    cif: cif,
    suratjalan: suratjalan,
  };

  localStorage.setItem("customerMappings", JSON.stringify(mappings));
  render();
  alert("Mapping berhasil disimpan!");
}

function render() {
  document.getElementById("mappingList").textContent = JSON.stringify(
    mappings,
    null,
    2
  );
}

function exportMapping() {
  const blob = new Blob([JSON.stringify(mappings, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "mapping.json";
  a.click();
}

function importMapping() {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".json";

  input.onchange = (e) => {
    const reader = new FileReader();
    reader.onload = () => {
      mappings = JSON.parse(reader.result);
      localStorage.setItem("customerMappings", JSON.stringify(mappings));
      render();
      alert("Berhasil import mapping!");
    };
    reader.readAsText(e.target.files[0]);
  };

  input.click();
}

load();
