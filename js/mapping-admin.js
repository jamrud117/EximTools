let mappings = {};

function load() {
  const saved = localStorage.getItem("customerMappings");
  if (saved) mappings = JSON.parse(saved);
  render();
}

function saveMapping() {
  const pt = document.getElementById("mapPT").value.trim();
  if (!pt) return alert("Nama customer wajib diisi!");

  mappings[pt] = {
    kode: document.getElementById("mapKode").value.trim(),
    uraian: document.getElementById("mapUraian").value.trim(),
    qty: document.getElementById("mapQty").value.trim(),
    cif: document.getElementById("mapCIF").value.trim(),
    suratjalan: document.getElementById("mapSJ").value.trim(),
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
