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
    copySubjectBtn: $("copySubjectBtn"),
    emailPreview: $("emailPreview"),
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

  const OUTPUT_COLUMNS = [
    "Tickets#",
    "Created on",
    "Department",
    "Prioritytitle",
    "Type",
    "subject",
    "Classification",
    "Organization",
    "Wing",
    "Closedon",
    "resolution_steps",
    "Status"
  ];

  function reorderColumns(row) {
    const result = {};
    OUTPUT_COLUMNS.forEach(column => {
      const key = Object.keys(row || {}).find(
        header => clean(header).toLowerCase() === column.toLowerCase()
      );
      result[column] = key ? row[key] : "";
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
        // Classification and Organization are generated fields. Ignore any
        // stale copies from the uploaded workbook so the processor remains
        // the single source of truth.
        if (/^(classification|organization)$/i.test(clean(header))) return;

        next[header] = clean(row[header]);
        if (header === subjectKey) {
          next.Classification = extractClassification(row[subjectKey]);
          next.Organization = getOrganization(row[subjectKey]);
        }
      });
      return reorderColumns(next);
    });
  }

  function normalizeStatus(value) {
    return clean(value).replace(/\s+/g, " ").toLowerCase();
  }

  function getColumnKey(row, target) {
    return Object.keys(row || {}).find(
      header => clean(header).toLowerCase() === target.toLowerCase()
    );
  }

  function summarize(rows) {
    const map = new Map();

    rows.forEach(row => {
      const client = clean(row.Classification) || "NABFID DC";
      const statusKey = getColumnKey(row, "status");
      if (!map.has(client)) {
        map.set(client, {
          Closed: 0,
          "Pending on COE": 0,
          "Pending On Customer": 0,
          total: 0
        });
      }

      const summary = map.get(client);
      const status = normalizeStatus(statusKey ? row[statusKey] : "");

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

    renderEmailPreview();
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

  function getCreatedOnValue(row) {
    const key = Object.keys(row || {}).find(
      header => clean(header).toLowerCase() === "created on"
    );
    return key ? row[key] : "";
  }

  function parseCreatedDateParts(value) {
    const text = clean(value);
    if (!text) return null;

    // Excel serial date/time.
    if (/^\d+(?:\.\d+)?$/.test(text)) {
      const serial = Number(text);
      if (serial > 0 && serial < 100000) {
        const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
        if (!Number.isNaN(date.getTime())) {
          return {
            day: date.getUTCDate(),
            month: date.getUTCMonth() + 1,
            year: date.getUTCFullYear()
          };
        }
      }
    }

    // Handles DD/MM/YYYY, DD-MM-YYYY, YYYY-MM-DD, and the same formats
    // when a time is appended.
    const match = text.match(/(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})/);
    if (match) {
      const a = Number(match[1]), b = Number(match[2]), c = Number(match[3]);
      let day, month, year;

      if (a >= 1000) {
        year = a; month = b; day = c;
      } else if (c >= 1000) {
        day = a; month = b; year = c;
      }

      if (
        Number.isInteger(day) &&
        Number.isInteger(month) &&
        Number.isInteger(year) &&
        day >= 1 && day <= 31 &&
        month >= 1 && month <= 12
      ) {
        return { day, month, year };
      }
    }

    // Last fallback for values such as "23 September 2026".
    const parsed = new Date(text);
    if (!Number.isNaN(parsed.getTime())) {
      return {
        day: parsed.getDate(),
        month: parsed.getMonth() + 1,
        year: parsed.getFullYear()
      };
    }

    return null;
  }

  function parseCreatedDate(value) {
    const parts = parseCreatedDateParts(value);
    if (!parts) return "";
    return String(parts.day).padStart(2, "0") +
      String(parts.month).padStart(2, "0") +
      String(parts.year).slice(-4);
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
    const parts = parseCreatedDateParts(getCreatedOnValue(processedRows[0]));
    const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sept","Oct","Nov","Dec"];
    const ordinal = day => {
      if (day % 100 >= 11 && day % 100 <= 13) return day + "th";
      return day + ({1:"st",2:"nd",3:"rd"}[day % 10] || "th");
    };
    const filename = parts
      ? "NaBFID Offenses " + ordinal(parts.day) + " " + months[parts.month - 1] + ".xlsx"
      : "NaBFID Offenses.xlsx";

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
            "text-align:center"
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

  function formatEmailDate(value) {
    const parts = parseCreatedDateParts(value);
    if (!parts) return "";

    const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
    const suffix = parts.day % 10 === 1 && parts.day !== 11 ? "st"
      : parts.day % 10 === 2 && parts.day !== 12 ? "nd"
      : parts.day % 10 === 3 && parts.day !== 13 ? "rd"
      : "th";

    return parts.day + suffix + " " + months[parts.month - 1] + " " + parts.year;
  }

  function getEmailClients(map) {
    // Keep the email client order identical to the visible summary table.
    const clients = [...map.keys()].map(clean).filter(Boolean).sort();
    return clients.length === 1
      ? clients[0]
      : clients.slice(0, -1).join(", ") + " and " + clients[clients.length - 1];
  }

  function buildEmailTable() {
    const summaryTable = document.querySelector("#summaryHead")?.closest("table");
    if (!summaryTable) return null;

    // Keep the email table identical to the locked Copy Table format.
    // Everything is inline so Outlook does not depend on the website CSS.
    const copiedTable = summaryTable.cloneNode(true);
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
          "text-align:center"
        ].join(";");
        cell.removeAttribute("class");
      });
    });

    return copiedTable;
  }

  function renderEmailPreview() {
    if (!els.emailPreview || !processedRows.length) return;

    const date = formatEmailDate(getCreatedOnValue(processedRows[0]));
    const map = summarize(processedRows);
    const clients = getEmailClients(map);
    const table = buildEmailTable();

    if (!date || !table) {
      els.emailPreview.innerHTML = '<div class="email-preview-empty">Unable to generate the email preview because the "Created on" date could not be determined.</div>';
      return;
    }

    const body = document.createElement("div");
    body.className = "email-preview-body";
    body.style.cssText = "max-width:920px;margin:0 auto;font-family:'Inter 18pt','Inter',Arial,sans-serif;font-size:10pt;color:#111827;line-height:1.5;";

    const subjectText = "NABFID Daily Offense Report || " + date;
    const subjectPreview = document.createElement("div");
    subjectPreview.className = "email-subject-preview";
    subjectPreview.dataset.subject = subjectText;
    subjectPreview.textContent = "Subject: " + subjectText;

    const greeting = document.createElement("p");
    greeting.textContent = "Hi Team,";
    greeting.style.cssText = "margin:0 0 18px 0;";

    const message = document.createElement("p");
    message.style.cssText = "margin:0 0 18px 0;";
    message.append("Please find the attached " + clients + " daily offense data for ");
    const strongDate = document.createElement("strong");
    strongDate.textContent = date;
    message.append(strongDate);
    message.append(".");

    body.append(greeting, message, table);
    els.emailPreview.replaceChildren(subjectPreview, body);
  }

  async function copySubject() {
    if (!processedRows.length) {
      alert("Please upload and process an Excel file first.");
      return;
    }

    const date = formatEmailDate(getCreatedOnValue(processedRows[0]));
    if (!date) {
      alert('Unable to determine the date from the "Created on" column.');
      return;
    }

    const subject = "NABFID Daily Offense Report || " + date;

    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(subject);
      } else {
        const area = document.createElement("textarea");
        area.value = subject;
        area.style.position = "fixed";
        area.style.left = "-9999px";
        document.body.appendChild(area);
        area.select();
        document.execCommand("copy");
        area.remove();
      }

      const oldLabel = els.copySubjectBtn.textContent;
      els.copySubjectBtn.textContent = "Subject Copied ✓";
      setTimeout(() => {
        els.copySubjectBtn.textContent = oldLabel;
      }, 1800);
    } catch (error) {
      console.error("Copy subject failed:", error);
      alert("Unable to copy the subject. Please try again.");
    }
  }

  async function copyEmail() {
    if (!processedRows.length) {
      alert("Please upload and process an Excel file first.");
      return;
    }

    renderEmailPreview();

    const previewBody = els.emailPreview?.querySelector(".email-preview-body");
    if (!previewBody) {
      alert('Unable to generate the email preview.');
      return;
    }

    const plainText = [...previewBody.querySelectorAll("p, tr")]
      .map(element => clean(element.textContent))
      .filter(Boolean)
      .join("\n");

    const html = previewBody.outerHTML;

    try {
      if (navigator.clipboard && window.ClipboardItem) {
        try {
          await navigator.clipboard.write([
            new ClipboardItem({
              "text/html": new Blob([html], { type: "text/html" }),
              "text/plain": new Blob([plainText], { type: "text/plain" })
            })
          ]);
        } catch (clipboardError) {
          const area = document.createElement("textarea");
          area.value = plainText;
          area.style.position = "fixed";
          area.style.left = "-9999px";
          document.body.appendChild(area);
          area.select();
          document.execCommand("copy");
          area.remove();
        }
      } else {
        const area = document.createElement("textarea");
        area.value = plainText;
        area.style.position = "fixed";
        area.style.left = "-9999px";
        document.body.appendChild(area);
        area.select();
        document.execCommand("copy");
        area.remove();
      }

      const oldLabel = els.copyEmailBtn.textContent;
      els.copyEmailBtn.textContent = "Email Copied ✓";
      setTimeout(() => {
        els.copyEmailBtn.textContent = oldLabel;
      }, 1800);
    } catch (error) {
      console.error("Copy email failed:", error);
      alert("Unable to copy the email. Please try again.");
    }
  }

  function recalculateDerivedFields() {
    processedRows.forEach(row => {
      const subjectKey = Object.keys(row).find(h => clean(h).toLowerCase() === "subject");
      if (!subjectKey) return;

      row.Classification = extractClassification(row[subjectKey]);
      row.Organization = getOrganization(row[subjectKey]);
    });
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
    els.copySubjectBtn.addEventListener("click", copySubject);

    els.saveBtn.addEventListener("click", () => {
      // Rebuild derived fields after edits so Classification and Organization
      // always reflect the current Subject value.
      recalculateDerivedFields();
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