const rowsBody = document.getElementById("rowsBody");
const addRowBtn = document.getElementById("addRowBtn");
const importBtn = document.getElementById("importBtn");
const importFileInput = document.getElementById("importFileInput");
const exportBtn = document.getElementById("exportBtn");
const appVersionEl = document.getElementById("appVersion");
const lastUpdatedEl = document.getElementById("lastUpdated");

const APP_VERSION = "v1.0.0";

let rowCounter = 0;

const fmt = new Intl.NumberFormat("tr-TR", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

function setBuildMeta() {
  if (appVersionEl) appVersionEl.textContent = APP_VERSION;
  if (lastUpdatedEl) lastUpdatedEl.textContent = new Date().toLocaleDateString("tr-TR");
}

function isValidHttpUrl(value) {
  try {
    const u = new URL(value);
    return u.protocol === "http:" || u.protocol === "https:";
  } catch (_err) {
    return false;
  }
}

async function generateQrDataUrl(url) {
  return QRCode.toDataURL(url, {
    width: 260,
    margin: 1,
    color: {
      dark: "#000000",
      light: "#FFFFFF",
    },
  });
}

function computeMetrics(rowEl) {
  const low = Number(rowEl.querySelector('[data-key="lowPrice"]').value || 0);
  const high = Number(rowEl.querySelector('[data-key="highPrice"]').value || 0);
  const avg = (low + high) / 2;
  const yearly = avg * 12;

  rowEl.querySelector('[data-view="avg"]').textContent = isFinite(avg) ? fmt.format(avg) : "0.00";
  rowEl.querySelector('[data-view="yearly"]').textContent = isFinite(yearly)
    ? fmt.format(yearly)
    : "0.00";
}

async function attachLink(rowEl, key) {
  const input = rowEl.querySelector(`[data-key="${key}"]`);
  const value = input.value.trim();
  if (!isValidHttpUrl(value)) {
    alert("Gecerli bir http/https linki girin.");
    return;
  }

  const qrBox = rowEl.querySelector(`[data-qr="${key}"]`);
  const qrDataUrl = await generateQrDataUrl(value);
  qrBox.innerHTML = `<img src="${qrDataUrl}" alt="QR" />`;
  qrBox.dataset.qrData = qrDataUrl;
}

async function setQrFromUrl(rowEl, key, url) {
  if (!isValidHttpUrl(url)) return;
  const qrBox = rowEl.querySelector(`[data-qr="${key}"]`);
  const qrDataUrl = await generateQrDataUrl(url);
  qrBox.innerHTML = `<img src="${qrDataUrl}" alt="QR" />`;
  qrBox.dataset.qrData = qrDataUrl;
}

function createInputCell(key, type = "text", placeholder = "") {
  return `<input type="${type}" step="0.01" data-key="${key}" placeholder="${placeholder}" />`;
}

function addRow(defaults = {}) {
  rowCounter += 1;
  const tr = document.createElement("tr");
  tr.dataset.rowId = String(rowCounter);

  tr.innerHTML = `
    <td>${createInputCell("device", "text", "Orn. Rigol DS1102Z-E")}</td>
    <td>${createInputCell("model", "text", "Model")}</td>
    <td>${createInputCell("lowPrice", "number", "0.00")}</td>
    <td>${createInputCell("lowSeller", "text", "Satici")}</td>
    <td>
      <div class="cell-stack">
        ${createInputCell("lowLink", "url", "https://...")}
        <button type="button" class="btn btn-link" data-link-btn="lowLink">Link Ekle + QR Uret</button>
      </div>
    </td>
    <td>${createInputCell("highPrice", "number", "0.00")}</td>
    <td>${createInputCell("highSeller", "text", "Satici")}</td>
    <td>
      <div class="cell-stack">
        ${createInputCell("highLink", "url", "https://...")}
        <button type="button" class="btn btn-link" data-link-btn="highLink">Link Ekle + QR Uret</button>
      </div>
    </td>
    <td><div class="metric" data-view="avg">0.00</div></td>
    <td><div class="metric" data-view="yearly">0.00</div></td>
    <td>
      <div class="qr-pair">
        <div class="qr-box" data-qr="lowLink">Dusuk Link QR</div>
        <div class="qr-box" data-qr="highLink">Yuksek Link QR</div>
      </div>
    </td>
    <td><button type="button" class="btn btn-delete">Sil</button></td>
  `;

  rowsBody.appendChild(tr);

  Object.entries(defaults).forEach(([k, v]) => {
    const el = tr.querySelector(`[data-key="${k}"]`);
    if (el) el.value = v;
  });

  // Imported/default link values should immediately show matching QR previews.
  const lowLinkValue = tr.querySelector('[data-key="lowLink"]').value.trim();
  const highLinkValue = tr.querySelector('[data-key="highLink"]').value.trim();
  if (isValidHttpUrl(lowLinkValue)) {
    setQrFromUrl(tr, "lowLink", lowLinkValue);
  }
  if (isValidHttpUrl(highLinkValue)) {
    setQrFromUrl(tr, "highLink", highLinkValue);
  }

  computeMetrics(tr);

  tr.addEventListener("input", (event) => {
    const target = event.target;
    if (target.matches('[data-key="lowPrice"], [data-key="highPrice"]')) {
      computeMetrics(tr);
    }

    if (target.matches('[data-key="lowLink"], [data-key="highLink"]')) {
      const key = target.dataset.key;
      const value = target.value.trim();
      const qrBox = tr.querySelector(`[data-qr="${key}"]`);
      if (isValidHttpUrl(value)) {
        setQrFromUrl(tr, key, value);
      } else {
        qrBox.textContent = key === "lowLink" ? "Dusuk Link QR" : "Yuksek Link QR";
        delete qrBox.dataset.qrData;
      }
    }
  });

  tr.querySelectorAll("[data-link-btn]").forEach((btn) => {
    btn.addEventListener("click", () => attachLink(tr, btn.dataset.linkBtn));
  });

  tr.querySelector(".btn-delete").addEventListener("click", () => {
    tr.remove();
  });
}

function getCellUrl(cell) {
  if (!cell || cell.value == null) return "";
  if (typeof cell.value === "object") {
    if (cell.value.hyperlink) return String(cell.value.hyperlink).trim();
    if (cell.value.text && isValidHttpUrl(String(cell.value.text))) return String(cell.value.text).trim();
  }
  if (typeof cell.value === "string" && isValidHttpUrl(cell.value.trim())) {
    return cell.value.trim();
  }
  return "";
}

function parseNumeric(value) {
  if (value == null || value === "") return 0;
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

async function importExcelFile(file) {
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const ws = workbook.worksheets[0];
  if (!ws) {
    alert("Excel dosyasinda sayfa bulunamadi.");
    return;
  }

  rowsBody.innerHTML = "";
  rowCounter = 0;

  for (let rowNum = 2; rowNum <= ws.rowCount; rowNum += 1) {
    const row = ws.getRow(rowNum);
    const device = String(row.getCell(1).value ?? "").trim();
    const model = String(row.getCell(2).value ?? "").trim();
    const lowLink = getCellUrl(row.getCell(5));
    const highLink = getCellUrl(row.getCell(8));
    const lowPrice = parseNumeric(row.getCell(3).value);
    const highPrice = parseNumeric(row.getCell(6).value);
    const lowSeller = String(row.getCell(4).value ?? "").trim();
    const highSeller = String(row.getCell(7).value ?? "").trim();

    const seemsDataRow =
      device || model || lowLink || highLink || lowPrice > 0 || highPrice > 0 || lowSeller || highSeller;

    if (!seemsDataRow) continue;

    addRow({
      device,
      model,
      lowPrice,
      lowSeller,
      lowLink,
      highPrice,
      highSeller,
      highLink,
    });
  }

  if (!rowsBody.querySelector("tr")) {
    addRow();
    alert("Veri satiri bulunamadi. Bos satir acildi.");
    return;
  }

}

function collectRows() {
  const rows = [];
  rowsBody.querySelectorAll("tr").forEach((tr) => {
    const rowData = {
      device: tr.querySelector('[data-key="device"]').value.trim(),
      model: tr.querySelector('[data-key="model"]').value.trim(),
      lowPrice: Number(tr.querySelector('[data-key="lowPrice"]').value || 0),
      lowSeller: tr.querySelector('[data-key="lowSeller"]').value.trim(),
      lowLink: tr.querySelector('[data-key="lowLink"]').value.trim(),
      highPrice: Number(tr.querySelector('[data-key="highPrice"]').value || 0),
      highSeller: tr.querySelector('[data-key="highSeller"]').value.trim(),
      highLink: tr.querySelector('[data-key="highLink"]').value.trim(),
      lowQr: tr.querySelector('[data-qr="lowLink"]').dataset.qrData || "",
      highQr: tr.querySelector('[data-qr="highLink"]').dataset.qrData || "",
    };

    if (rowData.device || rowData.model || rowData.lowLink || rowData.highLink) {
      rows.push(rowData);
    }
  });
  return rows;
}

async function exportExcel() {
  const rows = collectRows();
  if (!rows.length) {
    alert("Export icin en az bir satir girin.");
    return;
  }

  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet("Cihaz Fiyatlari");
  ws.views = [{ state: "frozen", ySplit: 1 }];

  ws.columns = [
    { header: "Cihaz Adi", key: "device", width: 22 },
    { header: "Model", key: "model", width: 15 },
    { header: "En Dusuk Fiyat (TL)", key: "lowPrice", width: 18 },
    { header: "Satis Yeri (Dusuk)", key: "lowSeller", width: 20 },
    { header: "Internet Sayfasi (Dusuk)", key: "lowLink", width: 36 },
    { header: "En Yuksek Fiyat (TL)", key: "highPrice", width: 18 },
    { header: "Satis Yeri (Yuksek)", key: "highSeller", width: 20 },
    { header: "Internet Sayfasi (Yuksek)", key: "highLink", width: 36 },
    { header: "Ortalama Fiyat", key: "avg", width: 14 },
    { header: "Adetli Toplam", key: "yearly", width: 14 },
  ];

  ws.getRow(1).height = 28;
  ws.getRow(1).eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF2F5D8A" },
    };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = {
      top: { style: "thin", color: { argb: "FFD0DAE6" } },
      left: { style: "thin", color: { argb: "FFD0DAE6" } },
      bottom: { style: "thin", color: { argb: "FFD0DAE6" } },
      right: { style: "thin", color: { argb: "FFD0DAE6" } },
    };
  });

  // Column accent bands provide a clearer scan in exported Excel.
  ["C", "F"].forEach((col) => {
    for (let r = 2; r <= rows.length + 1; r += 1) {
      ws.getCell(`${col}${r}`).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFF7FBFF" },
      };
    }
  });
  ["I", "J"].forEach((col) => {
    for (let r = 2; r <= rows.length + 1; r += 1) {
      ws.getCell(`${col}${r}`).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFEDF8F4" },
      };
    }
  });

  rows.forEach((r, index) => {
    const excelRow = index + 2;
    ws.getCell(`A${excelRow}`).value = r.device;
    ws.getCell(`B${excelRow}`).value = r.model;
    ws.getCell(`C${excelRow}`).value = r.lowPrice;
    ws.getCell(`D${excelRow}`).value = r.lowSeller;
    ws.getCell(`F${excelRow}`).value = r.highPrice;
    ws.getCell(`G${excelRow}`).value = r.highSeller;

    // Keep link cells visually clean in export; QR carries the navigation info.
    ws.getCell(`E${excelRow}`).value = "";
    ws.getCell(`H${excelRow}`).value = "";

    ws.getCell(`I${excelRow}`).value = { formula: `AVERAGE(C${excelRow},F${excelRow})` };
    ws.getCell(`J${excelRow}`).value = { formula: `I${excelRow}*12` };

    ws.getCell(`C${excelRow}`).numFmt = "#,##0.00";
    ws.getCell(`F${excelRow}`).numFmt = "#,##0.00";
    ws.getCell(`I${excelRow}`).numFmt = "#,##0.00";
    ws.getCell(`J${excelRow}`).numFmt = "#,##0.00";

    ws.getRow(excelRow).height = 190;

    for (let c = 1; c <= 10; c += 1) {
      const cell = ws.getCell(excelRow, c);
      if (excelRow % 2 === 0 && !cell.fill) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFDFEFF" },
        };
      }
      cell.border = {
        top: { style: "thin", color: { argb: "FFE1E6EE" } },
        left: { style: "thin", color: { argb: "FFE1E6EE" } },
        bottom: { style: "thin", color: { argb: "FFE1E6EE" } },
        right: { style: "thin", color: { argb: "FFE1E6EE" } },
      };
      cell.alignment = { vertical: "middle", wrapText: true };
    }

    ws.getCell(`A${excelRow}`).font = { bold: true, color: { argb: "FF1B2A41" } };

    if (r.lowQr) {
      const lowImageId = workbook.addImage({ base64: r.lowQr, extension: "png" });
      ws.addImage(lowImageId, {
        tl: { col: 4 + 0.15, row: excelRow - 1 + 0.05 },
        ext: { width: 232, height: 232 },
      });
    }

    if (r.highQr) {
      const highImageId = workbook.addImage({ base64: r.highQr, extension: "png" });
      ws.addImage(highImageId, {
        tl: { col: 7 + 0.15, row: excelRow - 1 + 0.05 },
        ext: { width: 232, height: 232 },
      });
    }
  });

  const noteRowIdx = rows.length + 2;
  const exportDate = new Date().toLocaleDateString("tr-TR");
  ws.mergeCells(`A${noteRowIdx}:J${noteRowIdx}`);
  ws.getCell(`A${noteRowIdx}`).value = `Not (${exportDate}): Fiyatlar pazar verisine gore degisiklik gosterebilir. Guncel fiyat icin ilgili internet sayfasini ziyaret edin.`;
  ws.getRow(noteRowIdx).height = 34;
  ws.getCell(`A${noteRowIdx}`).alignment = { wrapText: true, vertical: "middle" };
  ws.getCell(`A${noteRowIdx}`).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF1F5FB" },
  };
  ws.getCell(`A${noteRowIdx}`).font = { color: { argb: "FF41506A" }, italic: true };
  ws.getCell(`A${noteRowIdx}`).border = {
    top: { style: "thin", color: { argb: "FFD5DEEA" } },
    left: { style: "thin", color: { argb: "FFD5DEEA" } },
    bottom: { style: "thin", color: { argb: "FFD5DEEA" } },
    right: { style: "thin", color: { argb: "FFD5DEEA" } },
  };

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "Cihaz_Fiyat_Karsilastirmasi_web_export.xlsx";
  a.click();
  URL.revokeObjectURL(a.href);
}

addRowBtn.addEventListener("click", () => addRow());
importBtn.addEventListener("click", () => importFileInput.click());

importFileInput.addEventListener("change", async (event) => {
  const [file] = event.target.files || [];
  if (!file) return;

  try {
    await importExcelFile(file);
  } catch (_err) {
    alert("Excel okunurken hata olustu. Lutfen .xlsx formatinda dosya secin.");
  } finally {
    importFileInput.value = "";
  }
});

exportBtn.addEventListener("click", exportExcel);

setBuildMeta();

addRow({
  device: "Rigol DS1102Z-E",
  model: "DS1102Z-E",
  lowPrice: 16763.09,
  lowSeller: "Kartal Otomasyon",
  lowLink:
    "https://www.kartalotomasyon.com.tr/urun/rigol-ds1102z-e-100mhz-2-kanalli-1gs-s-dijital-osiloskop",
  highPrice: 41184.64,
  highSeller: "Perpa Otomasyon",
  highLink: "https://www.hepsiburada.com/rigol-ds1102z-e-100mhz-2-kanalli-1gs-s-dijital-osiloskop",
});
