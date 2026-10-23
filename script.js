
const $ = id => document.getElementById(id);

const templates = {
  winner: { image: null, data: null, positions: {} },
  participant: { image: null, data: null, positions: {} }
};

let participants = [];
let generatedFiles = [];
let currentType = "winner";
let placementMode = false;

$("readBtn").addEventListener("click", readExcel);
$("templateType").addEventListener("change", switchTemplate);
$("templateFile").addEventListener("change", loadTemplate);
$("placeModeBtn").addEventListener("click", beginPlacement);
$("templateCanvas").addEventListener("click", placeField);
$("fieldSelect").addEventListener("change", loadFieldStyle);
$("fontSize").addEventListener("change", saveFieldStyle);
$("fontSelect").addEventListener("change", saveFieldStyle);
$("boldText").addEventListener("change", saveFieldStyle);
$("textColor").addEventListener("change", saveFieldStyle);
$("clearFieldBtn").addEventListener("click", clearField);
$("clearAllBtn").addEventListener("click", clearAll);
$("previewBtn").addEventListener("click", previewFirst);
$("generateBtn").addEventListener("click", generatePDFs);
$("zipBtn").addEventListener("click", downloadZip);

window.addEventListener("resize", renderMarkers);

function status(id, message, error = false) {
  $(id).textContent = message;
  $(id).classList.toggle("error", error);
}

function currentTemplate() {
  return templates[currentType];
}

function switchTemplate() {
  currentType = $("templateType").value;
  placementMode = false;
  $("templateFile").value = "";

  const t = currentTemplate();

  if (t.image) {
    drawTemplate();
    status("templateStatus", `${currentType} template loaded.`);
  } else {
    $("templateCanvas").style.display = "none";
    $("emptyMessage").style.display = "block";
    $("markers").replaceChildren();
    status("templateStatus", `Upload the ${currentType} template image.`);
  }

  renderPositionList();
  loadFieldStyle();
}

function isWinner(position) {
  const p = String(position).trim().toLowerCase();

  return [
    "1", "1st", "first", "first place",
    "2", "2nd", "second", "second place",
    "3", "3rd", "third", "third place"
  ].includes(p);
}

function formatDate(value) {
  if (value instanceof Date && !isNaN(value)) {
    return `${String(value.getDate()).padStart(2, "0")}-${String(value.getMonth() + 1).padStart(2, "0")}-${value.getFullYear()}`;
  }

  return String(value ?? "").trim();
}

async function readExcel() {
  participants = [];
  generatedFiles = [];

  $("participantsTable").replaceChildren();
  $("downloads").replaceChildren();
  $("pdfPreview").replaceChildren();
  $("zipBtn").disabled = true;

  const file = $("excelFile").files[0];

  if (!file) {
    return status("excelStatus", "Choose an Excel file first.", true);
  }

  try {
    const workbook = XLSX.read(await file.arrayBuffer(), {
      type: "array",
      cellDates: true
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet, {
      defval: ""
    });

    if (!rows.length) {
      throw new Error("The spreadsheet is empty.");
    }

    if (rows.length > 5) {
      throw new Error("Maximum 5 participants allowed.");
    }

    const headers = Object.keys(rows[0]);

    const findColumn = name => headers.find(
      header => header.trim().toLowerCase() === name.toLowerCase()
    );

    const columns = ["Name", "Event", "Date", "Email", "Position"];

    const mapping = Object.fromEntries(
      columns.map(column => [column, findColumn(column)])
    );

    const missing = columns.filter(column => !mapping[column]);

    if (missing.length) {
      throw new Error("Missing columns: " + missing.join(", "));
    }

    participants = rows.map((row, index) => {
      const person = {};

      columns.forEach(column => {
        person[column] = column === "Date"
          ? formatDate(row[mapping[column]])
          : String(row[mapping[column]] ?? "").trim();
      });

      if (columns.some(column => !person[column])) {
        throw new Error(`Row ${index + 2}: fill all five columns.`);
      }

      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(person.Email)) {
        throw new Error(`Invalid email at row ${index + 2}.`);
      }

      person.type = isWinner(person.Position)
        ? "winner"
        : "participant";

      return person;
    });

    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");

    [...columns, "Certificate"].forEach(column => {
      const th = document.createElement("th");
      th.textContent = column;
      headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");

    participants.forEach(person => {
      const row = document.createElement("tr");

      [...columns.map(column => person[column]), person.type].forEach(value => {
        const td = document.createElement("td");
        td.textContent = value;
        row.appendChild(td);
      });

      tbody.appendChild(row);
    });

    table.appendChild(tbody);
    $("participantsTable").appendChild(table);

    status("excelStatus", `Successfully loaded ${participants.length} participant(s).`);
  } catch (error) {
    participants = [];
    status("excelStatus", error.message, true);
  }
}

function loadTemplate() {
  const file = $("templateFile").files[0];
  if (!file) return;

  if (!["image/jpeg", "image/png"].includes(file.type)) {
    return status("templateStatus", "Please upload a JPG, JPEG or PNG image.", true);
  }

  const reader = new FileReader();

  reader.onload = () => {
    const image = new Image();

    image.onload = () => {
      currentTemplate().image = image;
      currentTemplate().data = reader.result;

      drawTemplate();

      status(
        "templateStatus",
        `Loaded ${file.name} (${image.naturalWidth} × ${image.naturalHeight}).`
      );
    };

    image.onerror = () => {
      status("templateStatus", "Unable to load this image.", true);
    };

    image.src = reader.result;
  };

  reader.readAsDataURL(file);
}

function drawTemplate() {
  const template = currentTemplate();
  if (!template.image) return;

  const canvas = $("templateCanvas");

  canvas.width = template.image.naturalWidth;
  canvas.height = template.image.naturalHeight;

  canvas.getContext("2d").drawImage(template.image, 0, 0);

  canvas.style.display = "block";
  $("emptyMessage").style.display = "none";

  renderMarkers();
}

function beginPlacement() {
  if (!currentTemplate().image) {
    return status("editorStatus", "Upload a template image first.", true);
  }

  placementMode = true;

  status(
    "editorStatus",
    `Click the template where "${$("fieldSelect").value}" should begin.`
  );
}

function placeField(event) {
  if (!placementMode) return;

  const canvas = $("templateCanvas");
  const rect = canvas.getBoundingClientRect();

  const field = $("fieldSelect").value;

  currentTemplate().positions[field] = {
    x: Math.max(0, Math.min(1, (event.clientX - rect.left) / rect.width)),
    y: Math.max(0, Math.min(1, (event.clientY - rect.top) / rect.height)),
    size: fontSize(),
    font: $("fontSelect").value,
    bold: $("boldText").checked,
    color: $("textColor").value
  };

  placementMode = false;

  status("editorStatus", `${field} position saved for the ${currentType} template.`);

  renderMarkers();
  renderPositionList();
}

function fontSize() {
  return Math.max(6, Math.min(72, Number($("fontSize").value) || 24));
}

function saveFieldStyle() {
  const position = currentTemplate().positions[$("fieldSelect").value];
  if (!position) return;

  position.size = fontSize();
  position.font = $("fontSelect").value;
  position.bold = $("boldText").checked;
  position.color = $("textColor").value;

  renderPositionList();
}

function loadFieldStyle() {
  const position = currentTemplate().positions[$("fieldSelect").value];

  if (!position) {
    $("fontSize").value = 24;
    $("fontSelect").value = "helvetica";
    $("boldText").checked = false;
    $("textColor").value = "#111111";
    return;
  }

  $("fontSize").value = position.size;
  $("fontSelect").value = position.font;
  $("boldText").checked = position.bold;
  $("textColor").value = position.color;
}

function clearField() {
  delete currentTemplate().positions[$("fieldSelect").value];

  renderMarkers();
  renderPositionList();

  status("editorStatus", "Selected field position cleared.");
}

function clearAll() {
  currentTemplate().positions = {};

  renderMarkers();
  renderPositionList();

  status("editorStatus", `All positions cleared for the ${currentType} template.`);
}

function renderMarkers() {
  const holder = $("markers");
  holder.replaceChildren();

  Object.entries(currentTemplate().positions).forEach(([field, position]) => {
    const marker = document.createElement("div");

    marker.className = "marker";
    marker.style.left = `${position.x * 100}%`;
    marker.style.top = `${position.y * 100}%`;

    const label = document.createElement("span");
    label.textContent = field;

    marker.appendChild(label);
    holder.appendChild(marker);
  });
}

function renderPositionList() {
  const holder = $("positionList");
  holder.replaceChildren();

  const title = document.createElement("strong");
  title.textContent = `${currentType} template positions`;
  holder.appendChild(title);

  ["Name", "Event", "Date", "Position"].forEach(field => {
    const line = document.createElement("div");
    const position = currentTemplate().positions[field];

    line.textContent = position
      ? `${field}: placed · ${position.size} pt${position.bold ? " · bold" : ""}`
      : `${field}: not placed`;

    holder.appendChild(line);
  });
}

function previewFirst() {
  if (!participants.length) {
    return status("generateStatus", "Read the spreadsheet first.", true);
  }

  const person = participants[0];
  const template = templates[person.type];

  if (!template.image) {
    return status("generateStatus", `Upload the ${person.type} template first.`, true);
  }

  const missing = ["Name", "Event", "Date"].filter(
    field => !template.positions[field]
  );

  if (missing.length) {
    return status("generateStatus", `Place these fields first: ${missing.join(", ")}.`, true);
  }

  try {
    const blob = makePDF(person, template);
    const url = URL.createObjectURL(blob);

    const frame = document.createElement("iframe");
    frame.src = url;

    $("pdfPreview").replaceChildren(frame);
    status("generateStatus", `Previewing ${person.Name}'s certificate.`);
  } catch (error) {
    status("generateStatus", error.message, true);
  }
}

function hexToRgb(hex) {
  const value = hex.replace("#", "");

  return {
    r: parseInt(value.slice(0, 2), 16),
    g: parseInt(value.slice(2, 4), 16),
    b: parseInt(value.slice(4, 6), 16)
  };
}

function makePDF(person, template) {
  const { jsPDF } = window.jspdf;
  const image = template.image;

  const pdf = new jsPDF({
    orientation: image.naturalWidth >= image.naturalHeight ? "landscape" : "portrait",
    unit: "pt",
    format: "a4"
  });

  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();

  const scale = Math.min(
    pageW / image.naturalWidth,
    pageH / image.naturalHeight
  );

  const drawW = image.naturalWidth * scale;
  const drawH = image.naturalHeight * scale;

  const offsetX = (pageW - drawW) / 2;
  const offsetY = (pageH - drawH) / 2;

  pdf.addImage(
    template.data,
    template.data.startsWith("data:image/png") ? "PNG" : "JPEG",
    offsetX,
    offsetY,
    drawW,
    drawH
  );

  const values = {
    Name: person.Name,
    Event: person.Event,
    Date: person.Date,
    Position: person.Position
  };

  Object.entries(template.positions).forEach(([field, position]) => {
    if (values[field] === undefined) return;

    pdf.setFont(position.font, position.bold ? "bold" : "normal");
    pdf.setFontSize(position.size);

    const color = hexToRgb(position.color);
    pdf.setTextColor(color.r, color.g, color.b);

    const x = offsetX + position.x * drawW;
    const y = offsetY + position.y * drawH + position.size * 0.35;

    pdf.text(String(values[field]), x, y);
  });

  return pdf.output("blob");
}

async function generatePDFs() {
  if (!participants.length) {
    return status("generateStatus", "Read the spreadsheet first.", true);
  }

  generatedFiles = [];
  $("downloads").replaceChildren();
  $("zipBtn").disabled = true;

  status("generateStatus", "Generating certificates...");

  try {
    for (let index = 0; index < participants.length; index++) {
      const person = participants[index];
      const template = templates[person.type];

      if (!template.image) {
        throw new Error(`Upload the ${person.type} template first.`);
      }

      const requiredFields = ["Name", "Event", "Date"];
      const missing = requiredFields.filter(
        field => !template.positions[field]
      );

      if (missing.length) {
        throw new Error(
          `${person.type} template: place ${missing.join(", ")} first.`
        );
      }

      const safeName = person.Name.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");

      generatedFiles.push({
        person,
        blob: makePDF(person, template),
        filename: `${String(index + 1).padStart(2, "0")}_${safeName}_certificate.pdf`
      });
    }

    generatedFiles.forEach(file => {
      const link = document.createElement("a");

      link.className = "download-link";
      link.href = URL.createObjectURL(file.blob);
      link.download = file.filename;
      link.textContent =
        `${file.person.Name} — ${file.person.type} certificate · Download PDF`;

      $("downloads").appendChild(link);
    });

    $("zipBtn").disabled = false;

    status("generateStatus", `Successfully generated ${generatedFiles.length} PDF(s).`);
  } catch (error) {
    generatedFiles = [];
    status("generateStatus", error.message, true);
  }
}

async function downloadZip() {
  if (!generatedFiles.length) return;

  $("zipBtn").disabled = true;

  try {
    const zip = new JSZip();

    generatedFiles.forEach(file => {
      zip.file(file.filename, file.blob);
    });

    const blob = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");

    link.href = url;
    link.download = "certificates.zip";
    link.click();

    setTimeout(() => URL.revokeObjectURL(url), 1500);
  } catch (error) {
    status("generateStatus", "Could not create ZIP: " + error.message, true);
  } finally {
    $("zipBtn").disabled = false;
  }
}
