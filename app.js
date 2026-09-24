(function () {
  "use strict";

  const ORGANIZATIONS = {
    kpmg: "KPMG Assurance and Consulting Services LLP",
    default: "Edgeverve Systems Limited- NaBFID"
  };

  const EXCEL_FONT = "Inter 18pt";
  let processedRows = [];

  const $ = id => document.getElementById(id);

  const els = {
    fileInput: $("fileInput"),
    dropzone: $("dropzone"),
    fileName: $("fileName"),
    resultSection: $("resultSection"),
    previewBtn: $("previewBtn"),
    editBtn: $("editBtn"),
    downloadBtn: $("downloadBtn"),
    copyTableBtn: $("copyTableBtn"),
    copyEmailBtn: $("copyEmailBtn"),
    dataPanel: $("dataPanel"),
    dataTitle: $("dataTitle"),
    dataSubtitle: $("dataSubtitle"),
    dataHead: $("dataHead"),
    dataBody: $("dataBody"),
    saveBtn: $("saveBtn"),
    closePanelBtn: $("closePanelBtn")
  };

  function assertUI() {
    for (const [name, el] of Object.entries(els)) {
      if (!el) throw new Error("UI element missing: " + name);
    }
  }

  function clean(value) {
    return value == null ? "" : String(value).trim();
  }

  function extractClassification(subject) {
    const match = clean(subject).match(/Domain:\s*([^|]*)/i);
    return match ? clean(match[1]) || "NABFID DC" : "NABFID DC";
  }

  function getOrganization(subject) {
    return /KPMG/i.test(clean(subject))
      ? ORGANIZATIONS.kpmg
      : ORGANIZATIONS.default;
  }

  function reorderColumns(row, headers) {
    const columns = [];
    headers.forEach(header => {
      columns.push(header);
      if (header.toLowerCase() === "subject") {
        columns.push("Classification", "Organization");
      }
    });
    if (!columns.includes("Classification")) columns.push("Classification");
    if (!columns.includes("Organization")) columns.push("Organization");

    const result = {};
    [...new Set(columns)].forEach(column => {
      if (Object.prototype.hasOwnProperty.call(row, column)) result[column] = row[column];
    });
    return result;
  }

  function processSheet(rows) {
    if (!rows.length) throw new Error("The workbook is empty.");

    const headers = Object.keys(rows[0]);
    const subjectKey = headers.find(h => clean(h).toLowerCase() === "subject");
    if (!subjectKey) throw new Error('Required column "subject" was not found.');

    return rows.map(row => {
      const next = {};
      headers.forEach(header => {
        next[header] = clean(row[header]);
        if (header === subjectKey) {
          next.Classification = extractClassification(row[subjectKey]);
          next.Organization = getOrganization(row[subjectKey]);
        }
      });
      return reorderColumns(next, headers);
    });
  }

  function normalizeStatus(value) {
    return clean(value).toLowerCase();
  }

  function summarize(rows) {
    const map = new Map();

    rows.forEach(row => {
      const client = clean(row.Classification) || "NABFID DC";
      if (!map.has(client)) {
        map.set(client, {
          Closed: 0,
          "Pending on COE": 0,
          "Pending On Customer": 0,
          total: 0
        });
      }

      const summary = map.get(client);
      const status = normalizeStatus(row.Status);

      if (status === "closed") summary.Closed++;
      else if (status === "pending on coe") summary["Pending on COE"]++;
      else if (status === "pending on customer") summary["Pending On Customer"]++;
      summary.total++;
    });

    return map;
  }

  function renderSummary() {
    const map = summarize(processedRows);
    const head = $("summaryHead");
    const body = $("summaryBody");

    const headers = ["Client", "Closed", "Pending on COE", "Pending On Customer", "Grand Total"];
    head.innerHTML = "<tr>" + headers.map(h => "<th>" + h + "</th>").join("") + "</tr>";
    body.innerHTML = "";

    const totals = { Closed: 0, "Pending on COE": 0, "Pending On Customer": 0, total: 0 };

    [...map.keys()].sort().forEach(client => {
      const s = map.get(client);
      totals.Closed += s.Closed;
      totals["Pending on COE"] += s["Pending on COE"];
      totals["Pending On Customer"] += s["Pending On Customer"];
      totals.total += s.total;

      const tr = document.createElement("tr");
      [client, s.Closed, s["Pending on COE"], s["Pending On Customer"], s.total].forEach(v => {
        const td = document.createElement("td");
        td.textContent = v;
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });

    const totalRow = document.createElement("tr");
    ["Grand Total", totals.Closed, totals["Pending on COE"], totals["Pending On Customer"], totals.total].forEach(v => {
      const td = document.createElement("td");
      td.textContent = v;
      totalRow.appendChild(td);
    });
    body.appendChild(totalRow);

    $("statRecords").textContent = processedRows.length;
    $("statClients").textContent = map.size;
    $("statClosed").textContent = totals.Closed;
    $("statPending").textContent = totals["Pending on COE"] + totals["Pending On Customer"];
  }

  function renderData(editable) {
    if (!processedRows.length) return;

    els.dataPanel.classList.remove("hidden");
    els.saveBtn.classList.toggle("hidden", !editable);
    els.dataTitle.textContent = editable ? "Edit processed data" : "Preview processed data";
    els.dataSubtitle.textContent = editable
      ? "Edit cells below, then save edits to refresh the summary."
      : "Read-only preview of the generated Offenses sheet.";

    const headers = Object.keys(processedRows[0]);
    els.dataHead.innerHTML = "<tr>" + headers.map(h => "<th>" + h + "</th>").join("") + "</tr>";
    els.dataBody.innerHTML = "";

    processedRows.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");

      headers.forEach(header => {
        const td = document.createElement("td");
        td.className = "data-cell";

        if (editable) {
          const input = document.createElement("input");
          input.type = "text";
          input.value = row[header] || "";
          input.addEventListener("input", e => {
            processedRows[rowIndex][header] = e.target.value;
          });
          td.appendChild(input);
        } else {
          td.textContent = row[header] || "";
        }
        tr.appendChild(td);
      });

      els.dataBody.appendChild(tr);
    });
  }

  function parseCreatedDate(value) {
    const match = clean(value).match(/(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})/);
    if (!match) return "";

    const a = Number(match[1]), b = Number(match[2]), c = Number(match[3]);
    let day, month, year;

    if (a >= 1000) {
      year = a; month = b; day = c;
    } else if (c >= 1000) {
      day = a; month = b; year = c;
    } else {
      return "";
    }

    if (day < 1 || day > 31 || month < 1 || month > 12) return "";
    return String(day).padStart(2, "0") + String(month).padStart(2, "0") + String(year).slice(-4);
  }

  async function handleFile(file) {
    if (!file) return;

    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx", "xls"].includes(ext)) {
      alert("Please choose an Excel file (.xlsx or .xls).");
      return;
    }

    if (typeof XLSX === "undefined") {
      alert("Excel engine is not loaded. Please refresh the page and try again.");
      return;
    }

    els.fileName.textContent = "Processing: " + file.name;

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, {
        type: "array",
        cellDates: false,
        cellNF: true,
        cellText: true
      });

      if (!workbook.SheetNames || !workbook.SheetNames.length) {
        throw new Error("No sheets found in workbook.");
      }

      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      if (!sheet) throw new Error("Unable to read the first worksheet.");

      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
      processedRows = processSheet(rows);
      renderSummary();

      els.fileName.textContent = "Processed: " + file.name + " • " + processedRows.length + " records";
      els.resultSection.classList.remove("hidden");
      els.dataPanel.classList.add("hidden");

      requestAnimationFrame(() => els.resultSection.scrollIntoView({ behavior: "smooth", block: "start" }));
    } catch (error) {
      processedRows = [];
      console.error("Offense Processor V2:", error);
      els.fileName.textContent = "Processing failed: " + file.name;
      alert("Could not process file: " + (error && error.message ? error.message : error));
    }
  }

  function borders() {
    return {
      top: { style: "thin", color: { rgb: "202020" } },
      bottom: { style: "thin", color: { rgb: "202020" } },
      left: { style: "thin", color: { rgb: "202020" } },
      right: { style: "thin", color: { rgb: "202020" } }
    };
  }

  function headerStyle() {
    return {
      font: { name: EXCEL_FONT, sz: 10, family: 2, bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "0B2A5B" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: false },
      border: borders()
    };
  }

  function bodyStyle(horizontal) {
    return {
      font: { name: EXCEL_FONT, sz: 10, family: 2, color: { rgb: "374151" } },
      fill: { fgColor: { rgb: "FFFFFF" } },
      alignment: { horizontal: horizontal || "left", vertical: "center", wrapText: false },
      border: borders()
    };
  }

  function download() {
    if (!processedRows.length || typeof XLSX === "undefined") return;

    const workbook = XLSX.utils.book_new();
    const hStyle = headerStyle();

    const offenseSheet = XLSX.utils.json_to_sheet(processedRows);
    offenseSheet["!freeze"] = { xSplit: 0, ySplit: 1 };
    offenseSheet["!autofilter"] = { ref: offenseSheet["!ref"] };

    const headers = Object.keys(processedRows[0]);
    const widths = {
      "Tickets#": 16, "Created on": 20, "Department": 20, "Prioritytitle": 16,
      "Type": 16, "subject": 42, "Classification": 22, "Organization": 38,
      "Wing": 20, "Closedon": 20, "resolution_steps": 55, "Status": 18
    };
    offenseSheet["!cols"] = headers.map(h => ({ wch: widths[h] || 18 }));

    const range = XLSX.utils.decode_range(offenseSheet["!ref"]);
    for (let r = range.s.r; r <= range.e.r; r++) {
      offenseSheet["!rows"] = offenseSheet["!rows"] || [];
      offenseSheet["!rows"][r] = { hpt: r === 0 ? 24 : 20 };
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = offenseSheet[XLSX.utils.encode_cell({ r, c })];
        if (cell) cell.s = r === 0 ? hStyle : bodyStyle("left");
      }
    }
    XLSX.utils.book_append_sheet(workbook, offenseSheet, "Offenses");

    const map = summarize(processedRows);
    const summaryRows = [["Client", "Closed", "Pending on COE", "Pending On Customer", "Grand Total"]];
    const totals = [0, 0, 0, 0];

    [...map.keys()].sort().forEach(client => {
      const s = map.get(client);
      summaryRows.push([client, s.Closed, s["Pending on COE"], s["Pending On Customer"], s.total]);
      totals[0] += s.Closed;
      totals[1] += s["Pending on COE"];
      totals[2] += s["Pending On Customer"];
      totals[3] += s.total;
    });

    summaryRows.push(["Grand Total", totals[0], totals[1], totals[2], totals[3]]);

    const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
    summarySheet["!freeze"] = { xSplit: 0, ySplit: 1 };
    summarySheet["!autofilter"] = { ref: "A1:E" + summaryRows.length };
    summarySheet["!cols"] = [{ wch: 30 }, { wch: 14 }, { wch: 20 }, { wch: 24 }, { wch: 16 }];

    const sr = XLSX.utils.decode_range(summarySheet["!ref"]);
    for (let r = sr.s.r; r <= sr.e.r; r++) {
      summarySheet["!rows"] = summarySheet["!rows"] || [];
      summarySheet["!rows"][r] = { hpt: r === 0 ? 24 : 22 };

      for (let c = sr.s.c; c <= sr.e.c; c++) {
        const cell = summarySheet[XLSX.utils.encode_cell({ r, c })];
        if (!cell) continue;
        if (r === 0) cell.s = hStyle;
        else if (r === sr.e.r) {
          cell.s = {
            font: { name: EXCEL_FONT, sz: 10, family: 2, bold: true, color: { rgb: "FFFFFF" } },
            fill: { fgColor: { rgb: "0B2A5B" } },
            alignment: { horizontal: "center", vertical: "center", wrapText: false },
            border: borders()
          };
        } else {
          cell.s = bodyStyle(c === 0 ? "left" : "center");
        }
      }
    }

    XLSX.utils.book_append_sheet(workbook, summarySheet, "Classification");

    const output = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([output], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const date = parseCreatedDate(processedRows[0]["Created on"]);
    const filename = date ? "NABFID Offenses - " + date + ".xlsx" : "NABFID Offenses.xlsx";

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function copySummaryTable() {
    const table = document.querySelector("#summaryHead")?.closest("table");
    if (!table) return;

    const rows = [...table.querySelectorAll("tr")];
    const text = rows.map(row =>
      [...row.querySelectorAll("th,td")].map(cell => clean(cell.textContent)).join("\t")
    ).join("\n");

    try {
      // Build a self-contained table because pasted HTML does not inherit
      // the application's stylesheet.
      const copiedTable = table.cloneNode(true);
      copiedTable.removeAttribute("class");
      copiedTable.style.cssText = [
        "border-collapse:collapse",
        "border-spacing:0",
        "width:auto",
        "max-width:none",
        "table-layout:auto",
        "font-family:Inter,Arial,sans-serif",
        "font-size:9pt",
        "color:#111827"
      ].join(";");

      const copiedRows = [...copiedTable.querySelectorAll("tr")];
      copiedRows.forEach((row, rowIndex) => {
        const cells = [...row.querySelectorAll("th,td")];
        const isHeader = rowIndex === 0;
        const isGrandTotal = rowIndex === copiedRows.length - 1;

        cells.forEach((cell, cellIndex) => {
          cell.style.cssText = [
            "border:1px solid #202020",
            "padding:4px 8px",
            "line-height:1.35",
            "vertical-align:middle",
            "white-space:nowrap",
            "height:24px",
            isHeader || isGrandTotal ? "background:#0B2A5B" : "background:#FFFFFF",
            isHeader || isGrandTotal ? "color:#FFFFFF" : "color:#111827",
            isHeader || isGrandTotal ? "font-weight:700" : "font-weight:400",
            cellIndex === 0 ? "text-align:left" : "text-align:center"
          ].join(";");
          cell.removeAttribute("class");
        });
      });

      if (navigator.clipboard && window.ClipboardItem) {
        try {
          await navigator.clipboard.write([
            new ClipboardItem({
              "text/html": new Blob([copiedTable.outerHTML], { type: "text/html" }),
              "text/plain": new Blob([text], { type: "text/plain" })
            })
          ]);
        } catch (clipboardError) {
          const area = document.createElement("textarea");
          area.value = text;
          area.style.position = "fixed";
          area.style.left = "-9999px";
          document.body.appendChild(area);
          area.select();
          document.execCommand("copy");
          area.remove();
        }
      } else {
        const area = document.createElement("textarea");
        area.value = text;
        area.style.position = "fixed";
        area.style.left = "-9999px";
        document.body.appendChild(area);
        area.select();
        document.execCommand("copy");
        area.remove();
      }

      const oldLabel = els.copyTableBtn.textContent;
      els.copyTableBtn.textContent = "Copied ✓";
      setTimeout(() => {
        els.copyTableBtn.textContent = oldLabel;
      }, 1600);
    } catch (error) {
      console.error("Copy table failed:", error);
      alert("Unable to copy the table. Please try again.");
    }
  }

  function emailDate(value) {
    const m = clean(value).match(/(\\d{1,4})[\\/\\-.](\\d{1,2})[\\/\\-.](\\d{1,4})/);
    if (!m) return "";
    const a=Number(m[1]), b=Number(m[2]), c=Number(m[3]);
    const day = a >= 1000 ? c : a;
    const month = a >= 1000 ? b : b;
    const year = a >= 1000 ? a : c;
    if (day < 1 || day > 31 || month < 1 || month > 12 || year < 1000) return "";
    const months=["January","February","March","April","May","June","July","August","September","October","November","December"];
    const suffix = day%10===1 && day!==11 ? "st" : day%10===2 && day!==12 ? "nd" : day%10===3 && day!==13 ? "rd" : "th";
    return day + suffix + " " + months[month-1] + " " + year;
  }

  async function copyEmail() {
    if (!processedRows.length) return;
    const date = emailDate(processedRows[0]["Created on"]);
    if (!date) {
      alert('Unable to determine the date from "Created on".');
      return;
    }

    const map = summarize(processedRows);
    const total = { closed:0, coe:0, customer:0, all:0 };
    let body = "";
    [...map.keys()].sort().forEach(client => {
      const s=map.get(client);
      total.closed+=s.Closed;
      total.coe+=s["Pending on COE"];
      total.customer+=s["Pending On Customer"];
      total.all+=s.total;
      body += '<tr>' +
        '<td style="border:1px solid #202020;padding:4px 8px;white-space:nowrap;">'+client+'</td>' +
        '<td style="border:1px solid #202020;padding:4px 8px;text-align:center;">'+s.Closed+'</td>' +
        '<td style="border:1px solid #202020;padding:4px 8px;text-align:center;">'+s["Pending on COE"]+'</td>' +
        '<td style="border:1px solid #202020;padding:4px 8px;text-align:center;">'+s["Pending On Customer"]+'</td>' +
        '<td style="border:1px solid #202020;padding:4px 8px;text-align:center;">'+s.total+'</td></tr>';
    });

    const head='<tr>' +
      '<th style="border:1px solid #202020;padding:5px 9px;background:#0B2A5B;color:#fff;text-align:center;white-space:nowrap;">Client</th>' +
      '<th style="border:1px solid #202020;padding:5px 9px;background:#0B2A5B;color:#fff;text-align:center;white-space:nowrap;">Closed</th>' +
      '<th style="border:1px solid #202020;padding:5px 9px;background:#0B2A5B;color:#fff;text-align:center;white-space:nowrap;">Pending on COE</th>' +
      '<th style="border:1px solid #202020;padding:5px 9px;background:#0B2A5B;color:#fff;text-align:center;white-space:nowrap;">Pending On Customer</th>' +
      '<th style="border:1px solid #202020;padding:5px 9px;background:#0B2A5B;color:#fff;text-align:center;white-space:nowrap;">Grand Total</th></tr>';

    const grand='<tr>' +
      '<td style="border:1px solid #202020;padding:5px 9px;background:#0B2A5B;color:#fff;font-weight:700;">Grand Total</td>' +
      '<td style="border:1px solid #202020;padding:5px 9px;background:#0B2A5B;color:#fff;font-weight:700;text-align:center;">'+total.closed+'</td>' +
      '<td style="border:1px solid #202020;padding:5px 9px;background:#0B2A5B;color:#fff;font-weight:700;text-align:center;">'+total.coe+'</td>' +
      '<td style="border:1px solid #202020;padding:5px 9px;background:#0B2A5B;color:#fff;font-weight:700;text-align:center;">'+total.customer+'</td>' +
      '<td style="border:1px solid #202020;padding:5px 9px;background:#0B2A5B;color:#fff;font-weight:700;text-align:center;">'+total.all+'</td></tr>';

    const html='<div style="font-family:Arial,sans-serif;font-size:11pt;color:#111;line-height:1.5;">' +
      '<p style="margin:0 0 18px;">Hi Team,</p>' +
      '<p style="margin:0 0 18px;">Please find the attached NABFID DC, DR, OCI, KPMG and CCIL DC, DR daily offense data for <strong>'+date+'</strong>.</p>' +
      '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:auto;font-family:Arial,sans-serif;font-size:10pt;">'+head+'<tbody>'+body+grand+'</tbody></table>' +
      '<p style="margin:18px 0 0;">Thanks &amp; Regards,</p></div>';

    const text="Hi Team,\n\nPlease find the attached NABFID DC, DR, OCI, KPMG and CCIL DC, DR daily offense data for "+date+".\n\nThanks & Regards,";
    try {
      await navigator.clipboard.write([new ClipboardItem({
        "text/html": new Blob([html],{type:"text/html"}),
        "text/plain": new Blob([text],{type:"text/plain"})
      })]);
      const old=els.copyEmailBtn.textContent;
      els.copyEmailBtn.textContent="Email Copied ✓";
      setTimeout(()=>els.copyEmailBtn.textContent=old,1800);
    } catch (error) {
      console.error(error);
      alert("Unable to copy the email. Please try again.");
    }
  }

  function showDataPanel(editable) {
    if (!processedRows.length) {
      alert("Please upload and process an Excel file first.");
      return;
    }
    renderData(editable);
    els.dataPanel.classList.remove("hidden");
    requestAnimationFrame(() => els.dataPanel.scrollIntoView({ behavior: "smooth", block: "start" }));
  }

  function bindEvents() {
    els.fileInput.addEventListener("change", event => {
      const file = event.target.files && event.target.files[0];
      if (file) handleFile(file);
    });

    els.dropzone.addEventListener("dragover", event => {
      event.preventDefault();
      els.dropzone.classList.add("drop-active");
    });

    els.dropzone.addEventListener("dragleave", () => {
      els.dropzone.classList.remove("drop-active");
    });

    els.dropzone.addEventListener("drop", event => {
      event.preventDefault();
      els.dropzone.classList.remove("drop-active");
      const file = event.dataTransfer.files && event.dataTransfer.files[0];
      if (file) handleFile(file);
    });

    els.previewBtn.addEventListener("click", () => showDataPanel(false));
    els.editBtn.addEventListener("click", () => showDataPanel(true));
    els.downloadBtn.addEventListener("click", download);
    els.copyTableBtn.addEventListener("click", copySummaryTable);
    els.copyEmailBtn.addEventListener("click", copyEmail);

    els.saveBtn.addEventListener("click", () => {
      renderSummary();
      renderData(true);
    });

    els.closePanelBtn.addEventListener("click", () => {
      els.dataPanel.classList.add("hidden");
    });
  }

  function init() {
    try {
      assertUI();
      bindEvents();
      if (typeof XLSX === "undefined") {
        els.fileName.textContent = "Excel engine unavailable — refresh the page";
        console.error("Offense Processor V2: XLSX library is not loaded.");
      }
    } catch (error) {
      console.error("Offense Processor V2 initialization failed:", error);
      alert("The application could not initialize. Please refresh the page.");
    }
  }

  init();
})();