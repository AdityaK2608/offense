(function () {
  "use strict";

  const ORGANIZATIONS = {
    kpmg: "KPMG Assurance and Consulting Services LLP",
    default: "Edgeverve Systems Limited- NaBFID"
  };
  const OUTPUT_COLUMNS = [
    "Tickets#", "Created on", "Department", "Prioritytitle", "Type", "subject",
    "Classification", "Organization", "Wing", "Closedon", "resolution_steps", "Status"
  ];
  const COLUMN_WIDTHS = {
    "Tickets#": 16, "Created on": 20, "Department": 20, "Prioritytitle": 16,
    "Type": 16, "subject": 42, "Classification": 22, "Organization": 38,
    "Wing": 20, "Closedon": 20, "resolution_steps": 55, "Status": 18
  };
  const SUMMARY_HEADERS = ["Client", "Closed", "Pending on COE", "Pending On Customer", "Grand Total"];
  const EXCEL_FONT = "Inter";
  const EXCLUDED_SUBJECT_PATTERNS = [
    /not\s+reporting/i,
    /devices\s+not\s+working/i,
    /devices\s+not\s+reporting/i
  ];
  const $ = id => document.getElementById(id);
  let rawRows = [];
  let processedRows = [];
  let selectedReportDate = "";
  let availableReportDates = [];
  let calendarCursor = new Date();

  const els = {
    fileInput: $("fileInput"), dropzone: $("dropzone"), fileName: $("fileName"),
    dateSelection: $("dateSelection"), reportDate: $("reportDate"), datePickerButton: $("datePickerButton"), datePickerValue: $("datePickerValue"),
    datePickerPopover: $("datePickerPopover"), calendarPrev: $("calendarPrev"), calendarNext: $("calendarNext"),
    calendarMonth: $("calendarMonth"), calendarHint: $("calendarHint"), calendarGrid: $("calendarGrid"), calendarToday: $("calendarToday"),
    processBtn: $("processBtn"),
    dateSelectionStatus: $("dateSelectionStatus"), resultSection: $("resultSection"), previewBtn: $("previewBtn"), editBtn: $("editBtn"),
    copyTableBtn: $("copyTableBtn"), downloadBtn: $("downloadBtn"),
    copyEmailBtn: $("copyEmailBtn"), copySubjectBtn: $("copySubjectBtn"),
    emailPreview: $("emailPreview"), dataPanel: $("dataPanel"), dataTitle: $("dataTitle"),
    dataSubtitle: $("dataSubtitle"), dataHead: $("dataHead"), dataBody: $("dataBody"),
    saveBtn: $("saveBtn"), closePanelBtn: $("closePanelBtn")
  };

  function assertUI() {
    for (const [name, el] of Object.entries(els)) {
      if (!el) throw new Error("UI element missing: " + name);
    }
  }

  function clean(value) {
    return value == null ? "" : String(value).trim();
  }

  function sourceValue(row, column) {
    const key = Object.keys(row || {}).find(
      h => clean(h).toLowerCase() === column.toLowerCase()
    );
    return key ? clean(row[key]) : "";
  }

  const SOURCE_DATE_FORMAT = "d-m-yyyy h:mm:ss AM/PM";

  function toExcelDateValue(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    const text = clean(value);
    if (!text) return "";

    if (/^\d+(?:\.\d+)?$/.test(text)) {
      const serial = Number(text);
      if (serial > 0 && serial < 100000 && typeof XLSX !== "undefined" && XLSX.SSF?.parse_date_code) {
        const parsed = XLSX.SSF.parse_date_code(serial);
        if (parsed?.y && parsed?.m && parsed?.d) {
          return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H || 0, parsed.M || 0, parsed.S || 0);
        }
      }
    }

    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? text : parsed;
  }

  function formatDateForDisplay(value) {
    const date = toExcelDateValue(value);
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return clean(value);
    let hours = date.getHours();
    const suffix = hours >= 12 ? "PM" : "AM";
    hours = hours % 12 || 12;
    return date.getDate() + "-" + (date.getMonth() + 1) + "-" + date.getFullYear() +
      " " + hours + ":" + String(date.getMinutes()).padStart(2, "0") + ":" + String(date.getSeconds()).padStart(2, "0") +
      " " + suffix;
  }

  function extractClassification(subject) {
    const match = clean(subject).match(/Domain:\s*([^|]*)/i);
    return match ? clean(match[1]) || "NABFID DC" : "NABFID DC";
  }

  function getOrganization(subject) {
    return /KPMG/i.test(clean(subject)) ? ORGANIZATIONS.kpmg : ORGANIZATIONS.default;
  }

  function normalizeRule(value) {
    return clean(value)
      .replace(/\s+/g, " ")
      .replace(/\s*-\s*/g, "-")
      .replace(/\s*_\s*/g, "_")
      .toLowerCase();
  }

  function extractRuleName(row) {
    // RuleName is always embedded in the subject:
    // "Domain: ...|Offence_ID: ...|RuleName: <rule>"
    const subject = sourceValue(row, "subject");
    const match = subject.match(/(?:^|\|)\s*RuleName\s*:\s*([^|]*)/i);
    return match ? clean(match[1]) : "";
  }

  function isNabfidRecord(row) {
    return /nabfid/i.test(sourceValue(row, "subject")) || /nabfid/i.test(sourceValue(row, "Organization"));
  }

  function isExcludedNabfidRule(row) {
    const subject = sourceValue(row, "subject");
    if (!subject) return false;

    // Exclude any NaBFID subject containing these operational phrases.
    return EXCLUDED_SUBJECT_PATTERNS.some(pattern => pattern.test(subject));
  }

  function isBlankResolution(row) {
    return !clean(sourceValue(row, "resolution_steps"));
  }

  function isSocAlert(row) {
    return clean(sourceValue(row, "Type")).toLowerCase() === "soc alert";
  }

  function processSheet(rows) {
    if (!rows.length) throw new Error("The workbook is empty.");
    if (!Object.keys(rows[0]).some(h => clean(h).toLowerCase() === "subject")) {
      throw new Error('Required column "subject" was not found.');
    }

    return rows.map(row => {
      const subject = sourceValue(row, "subject");
      return {
        "Tickets#": sourceValue(row, "Tickets#"),
        "Created on": toExcelDateValue(sourceValue(row, "Created on")),
        "Department": sourceValue(row, "Department"),
        "Prioritytitle": sourceValue(row, "Prioritytitle"),
        "Type": "SOC Alert",
        subject,
        "Classification": extractClassification(subject),
        "Organization": getOrganization(subject),
        "Wing": sourceValue(row, "Wing"),
        "Closedon": toExcelDateValue(sourceValue(row, "Closedon")),
        "resolution_steps": sourceValue(row, "resolution_steps"),
        "Status": sourceValue(row, "Status")
      };
    });
  }

  function normalizeStatus(value) {
    return clean(value).replace(/\s+/g, " ").toLowerCase();
  }

  function summarize(rows) {
    const map = new Map();
    rows.forEach(row => {
      const client = clean(row.Classification) || "NABFID DC";
      if (!map.has(client)) {
        map.set(client, { Closed: 0, "Pending on COE": 0, "Pending On Customer": 0, total: 0 });
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

  function appendCells(row, values) {
    values.forEach(value => {
      const td = document.createElement("td");
      td.textContent = value;
      row.appendChild(td);
    });
  }

  function renderSummary() {
    const map = summarize(processedRows);
    const head = $("summaryHead");
    const body = $("summaryBody");
    head.innerHTML = "<tr>" + SUMMARY_HEADERS.map(h => "<th>" + h + "</th>").join("") + "</tr>";
    body.innerHTML = "";

    const totals = { Closed: 0, "Pending on COE": 0, "Pending On Customer": 0, total: 0 };
    [...map.keys()].sort().forEach(client => {
      const s = map.get(client);
      totals.Closed += s.Closed;
      totals["Pending on COE"] += s["Pending on COE"];
      totals["Pending On Customer"] += s["Pending On Customer"];
      totals.total += s.total;
      const tr = document.createElement("tr");
      appendCells(tr, [client, s.Closed, s["Pending on COE"], s["Pending On Customer"], s.total]);
      body.appendChild(tr);
    });

    const totalRow = document.createElement("tr");
    appendCells(totalRow, ["Grand Total", totals.Closed, totals["Pending on COE"], totals["Pending On Customer"], totals.total]);
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

    els.dataHead.innerHTML = "<tr>" + OUTPUT_COLUMNS.map(h => "<th>" + h + "</th>").join("") + "</tr>";
    els.dataBody.innerHTML = "";

    processedRows.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");
      OUTPUT_COLUMNS.forEach(header => {
        const td = document.createElement("td");
        if (editable) {
          const input = document.createElement("input");
          input.type = "text";
          input.value = row[header] || "";
          input.addEventListener("input", e => { processedRows[rowIndex][header] = e.target.value; });
          td.appendChild(input);
        } else {
          td.textContent = (header === "Created on" || header === "Closedon") ? formatDateForDisplay(row[header]) : (row[header] || "");
        }
        tr.appendChild(td);
      });
      els.dataBody.appendChild(tr);
    });
  }

  function getCreatedOnValue(row) {
    return row?.["Created on"] || "";
  }

  function parseCreatedDateParts(value) {
    const text = clean(value);
    if (!text) return null;

    if (/^\d+(?:\.\d+)?$/.test(text)) {
      const serial = Number(text);
      if (serial > 0 && serial < 100000) {
        const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
        if (!Number.isNaN(date.getTime())) {
          return { day: date.getUTCDate(), month: date.getUTCMonth() + 1, year: date.getUTCFullYear() };
        }
      }
    }

    const match = text.match(/(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})/);
    if (match) {
      const a = Number(match[1]), b = Number(match[2]), c = Number(match[3]);
      const [day, month, year] = a >= 1000 ? [c, b, a] : [a, b, c];
      if (year >= 1000 && day >= 1 && day <= 31 && month >= 1 && month <= 12) {
        return { day, month, year };
      }
    }

    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime())
      ? null
      : { day: parsed.getDate(), month: parsed.getMonth() + 1, year: parsed.getFullYear() };
  }

  function ordinal(day) {
    if (day % 100 >= 11 && day % 100 <= 13) return day + "th";
    return day + ({ 1: "st", 2: "nd", 3: "rd" }[day % 10] || "th");
  }

  function formatEmailDate(value) {
    const parts = parseCreatedDateParts(value);
    if (!parts) return "";
    const months = ["January", "February", "March", "April", "May", "June", "July", "August",
      "September", "October", "November", "December"];
    return ordinal(parts.day) + " " + months[parts.month - 1] + " " + parts.year;
  }

  function copyFallback(text) {
    const area = document.createElement("textarea");
    area.value = text;
    area.style.cssText = "position:fixed;left:-9999px;";
    document.body.appendChild(area);
    area.select();
    document.execCommand("copy");
    area.remove();
  }

  async function copyText(text) {
    if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(text);
    else copyFallback(text);
  }

  async function copyRich(html, text) {
    if (navigator.clipboard && window.ClipboardItem) {
      try {
        await navigator.clipboard.write([
          new ClipboardItem({
            "text/html": new Blob([html], { type: "text/html" }),
            "text/plain": new Blob([text], { type: "text/plain" })
          })
        ]);
        return;
      } catch (_) {}
    }
    copyFallback(text);
  }

  function copiedTableStyle(table) {
    table.removeAttribute("class");
    table.style.cssText = [
      "border-collapse:collapse", "border-spacing:0", "width:auto", "max-width:none",
      "table-layout:auto", "font-family:\"Inter 18pt\",Inter,Arial,sans-serif", "font-size:10pt", "color:#111827"
    ].join(";");

    [...table.querySelectorAll("tr")].forEach((row, rowIndex, rows) => {
      const special = rowIndex === 0 || rowIndex === rows.length - 1;
      row.querySelectorAll("th,td").forEach(cell => {
        cell.style.cssText = [
          "border:1px solid #202020", "padding:4px 8px", "line-height:1.35", "font-family:\"Inter 18pt\",Inter,Arial,sans-serif", "font-size:10pt",
          "vertical-align:middle", "white-space:nowrap", "height:24px",
          special ? "background:#0B2A5B" : "background:#FFFFFF",
          special ? "color:#FFFFFF" : "color:#111827",
          special ? "font-weight:700" : "font-weight:400", "text-align:center"
        ].join(";");
        cell.removeAttribute("class");
      });
    });
    return table;
  }

  function summaryTable() {
    return document.querySelector("#summaryHead")?.closest("table");
  }

  async function copySummaryTable() {
    const table = summaryTable();
    if (!table) return;
    const rows = [...table.querySelectorAll("tr")];
    const text = rows.map(row => [...row.querySelectorAll("th,td")].map(cell => clean(cell.textContent)).join("\t")).join("\n");

    try {
      const copied = copiedTableStyle(table.cloneNode(true));
      await copyRich(copied.outerHTML, text);
      showCopied(els.copyTableBtn, "Copied ✓", 1600);
    } catch (error) {
      console.error("Copy table failed:", error);
      alert("Unable to copy the table. Please try again.");
    }
  }

  function showCopied(button, label, duration = 1800) {
    const oldLabel = button.textContent;
    button.textContent = label;
    setTimeout(() => { button.textContent = oldLabel; }, duration);
  }

  function getEmailClients(map) {
    const clients = [...map.keys()].map(clean).filter(Boolean).sort();
    if (clients.length <= 1) return clients[0] || "";
    return clients.slice(0, -1).join(", ") + " and " + clients.at(-1);
  }

  function buildEmailTable() {
    const table = summaryTable();
    return table ? copiedTableStyle(table.cloneNode(true)) : null;
  }

  function renderEmailPreview() {
    if (!processedRows.length) return;
    const date = formatEmailDate(getCreatedOnValue(processedRows[0]));
    const table = buildEmailTable();
    if (!date || !table) {
      els.emailPreview.innerHTML = '<div class="email-preview-empty">Unable to generate the email preview because the "Created on" date could not be determined.</div>';
      return;
    }

    const clients = getEmailClients(summarize(processedRows));
    const subject = "NABFID Daily Offense Report || " + date;
    const subjectPreview = document.createElement("div");
    subjectPreview.className = "email-subject-preview";
    subjectPreview.textContent = "Subject: " + subject;

    const body = document.createElement("div");
    body.className = "email-preview-body";
    body.style.cssText = "max-width:920px;margin:0 auto;font-family:\"Inter 18pt\",Inter,Arial,sans-serif;font-size:10pt;color:#111827;line-height:1.5;"

    const greeting = document.createElement("p");
    greeting.textContent = "Hi Team,";
    greeting.style.cssText = "margin:0 0 18px 0;font-family:\"Inter 18pt\",Inter,Arial,sans-serif;font-size:10pt;color:#111827;"

    const message = document.createElement("p");
    message.style.cssText = "margin:0 0 18px 0;font-family:\"Inter 18pt\",Inter,Arial,sans-serif;font-size:10pt;color:#111827;";
    message.append("Please find the attached " + clients + " daily offense data for ");
    const strongDate = document.createElement("strong");
    strongDate.textContent = date;
    strongDate.style.cssText = "font-family:\"Inter 18pt\",Inter,Arial,sans-serif;font-size:10pt;color:#111827;font-weight:700;";
    message.append(strongDate, ".");

    body.append(greeting, message, table);
    els.emailPreview.replaceChildren(subjectPreview, body);
  }

  async function copySubject() {
    if (!processedRows.length) return alert("Please upload and process an Excel file first.");
    const date = formatEmailDate(getCreatedOnValue(processedRows[0]));
    if (!date) return alert('Unable to determine the date from the "Created on" column.');

    try {
      await copyText("NABFID Daily Offense Report || " + date);
      showCopied(els.copySubjectBtn, "Subject Copied ✓");
    } catch (error) {
      console.error("Copy subject failed:", error);
      alert("Unable to copy the subject. Please try again.");
    }
  }

  async function copyEmail() {
    if (!processedRows.length) return alert("Please upload and process an Excel file first.");
    renderEmailPreview();

    const previewBody = els.emailPreview?.querySelector(".email-preview-body");
    if (!previewBody) return alert("Unable to generate the email preview.");

    const plainText = [...previewBody.querySelectorAll("p, tr")]
      .map(element => clean(element.textContent)).filter(Boolean).join("\n");

    try {
      await copyRich(previewBody.outerHTML, plainText);
      showCopied(els.copyEmailBtn, "Email Copied ✓");
    } catch (error) {
      console.error("Copy email failed:", error);
      alert("Unable to copy the email. Please try again.");
    }
  }

  function borders() {
    const side = { style: "thin", color: { rgb: "202020" } };
    return { top: side, bottom: side, left: side, right: side };
  }

  function headerStyle() {
    return {
      font: { name: EXCEL_FONT, sz: 10, family: 2, bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "0B2A5B" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: false },
      border: borders()
    };
  }

  function bodyStyle(horizontal = "left", bold = false) {
    return {
      font: { name: EXCEL_FONT, sz: 10, family: 2, bold, color: { rgb: bold ? "FFFFFF" : "374151" } },
      fill: { fgColor: { rgb: bold ? "0B2A5B" : "FFFFFF" } },
      alignment: { horizontal, vertical: "center", wrapText: false },
      border: borders()
    };
  }

  function styleSheet(sheet, headerRowHeight, bodyRowHeight, totalRow = -1) {
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    sheet["!rows"] = [];
    const hStyle = headerStyle();

    for (let r = range.s.r; r <= range.e.r; r++) {
      sheet["!rows"][r] = { hpt: r === 0 ? headerRowHeight : bodyRowHeight };
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c })];
        if (!cell) continue;
        cell.s = r === 0 ? hStyle : r === totalRow ? bodyStyle("center", true) : bodyStyle(c === 0 ? "left" : "center");
      }
    }
  }

  function download() {
    if (!processedRows.length || typeof XLSX === "undefined") return;

    const workbook = XLSX.utils.book_new();
    const offenseSheet = XLSX.utils.json_to_sheet(processedRows);
    ["B", "J"].forEach(col => {
      for (let r = 2; r <= processedRows.length + 1; r++) {
        const cell = offenseSheet[col + r];
        if (cell && cell.v instanceof Date) cell.z = SOURCE_DATE_FORMAT;
      }
    });
    offenseSheet["!freeze"] = { xSplit: 0, ySplit: 1 };
    offenseSheet["!autofilter"] = { ref: offenseSheet["!ref"] };
    offenseSheet["!cols"] = OUTPUT_COLUMNS.map(h => ({ wch: COLUMN_WIDTHS[h] || 18 }));
    styleSheet(offenseSheet, 24, 20);
    XLSX.utils.book_append_sheet(workbook, offenseSheet, "Offenses");

    const map = summarize(processedRows);
    const rows = [SUMMARY_HEADERS];
    const totals = [0, 0, 0, 0];

    [...map.keys()].sort().forEach(client => {
      const s = map.get(client);
      rows.push([client, s.Closed, s["Pending on COE"], s["Pending On Customer"], s.total]);
      totals[0] += s.Closed;
      totals[1] += s["Pending on COE"];
      totals[2] += s["Pending On Customer"];
      totals[3] += s.total;
    });
    rows.push(["Grand Total", ...totals]);

    const summarySheet = XLSX.utils.aoa_to_sheet(rows);
    summarySheet["!freeze"] = { xSplit: 0, ySplit: 1 };
    summarySheet["!autofilter"] = { ref: "A1:E" + rows.length };
    summarySheet["!cols"] = [{ wch: 30 }, { wch: 14 }, { wch: 20 }, { wch: 24 }, { wch: 16 }];
    styleSheet(summarySheet, 24, 22, rows.length - 1);
    XLSX.utils.book_append_sheet(workbook, summarySheet, "Classification");

    const output = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([output], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const parts = parseCreatedDateParts(getCreatedOnValue(processedRows[0]));
    const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sept","Oct","Nov","Dec"];
    const filename = parts ? "NaBFID Offenses " + ordinal(parts.day) + " " + months[parts.month - 1] + ".xlsx" : "NaBFID Offenses.xlsx";

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function recalculateDerivedFields() {
    processedRows.forEach(row => {
      row.Classification = extractClassification(row.subject);
      row.Organization = getOrganization(row.subject);
    });
  }

  function showDataPanel(editable) {
    if (!processedRows.length) return alert("Please upload and process an Excel file first.");
    renderData(editable);
    requestAnimationFrame(() => els.dataPanel.scrollIntoView({ behavior: "smooth", block: "start" }));
  }

  function dateKeyFromValue(value) {
    const parts = parseCreatedDateParts(value);
    return parts ? [parts.year, String(parts.month).padStart(2, "0"), String(parts.day).padStart(2, "0")].join("-") : "";
  }

  function resetForNewUpload() {
    rawRows = [];
    processedRows = [];
    selectedReportDate = "";
    availableReportDates = [];
    els.reportDate.value = "";
    els.datePickerValue.textContent = "Select a date";
    els.datePickerButton.setAttribute("aria-expanded", "false");
    els.datePickerPopover.classList.add("hidden");
    els.processBtn.disabled = true;
    els.dateSelectionStatus.textContent = "";
    els.dateSelection.classList.add("hidden");
    els.resultSection.classList.add("hidden");
    els.dataPanel.classList.add("hidden");
  }

  function showDateSelection(rows) {
    const dates = [...new Set(rows.map(row => dateKeyFromValue(sourceValue(row, "Created on"))).filter(Boolean))].sort();
    if (!dates.length) throw new Error('No valid dates were found in the "Created on" column.');
    availableReportDates = dates;
    calendarCursor = new Date(Number(dates[0].slice(0, 4)), Number(dates[0].slice(5, 7)) - 1, 1);
    els.reportDate.value = "";
    els.datePickerValue.textContent = "Select a date";
    els.processBtn.disabled = true;
    els.dateSelectionStatus.textContent = dates.length + " date" + (dates.length === 1 ? "" : "s") + " available in the workbook.";
    renderCalendar();
    els.dateSelection.classList.remove("hidden");
  }

  function formatPickerDate(key) {
    const [year, month, day] = key.split("-").map(Number);
    const date = new Date(year, month - 1, day);
    return date.toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" });
  }

  function renderCalendar() {
    const year = calendarCursor.getFullYear();
    const month = calendarCursor.getMonth();
    const monthName = calendarCursor.toLocaleDateString("en-US", { month: "long", year: "numeric" });
    els.calendarMonth.textContent = monthName;
    els.calendarHint.textContent = availableReportDates.length + " date" + (availableReportDates.length === 1 ? "" : "s") + " available";
    els.calendarGrid.innerHTML = "";

    const firstDay = new Date(year, month, 1);
    const startOffset = (firstDay.getDay() + 6) % 7;
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const prevDays = new Date(year, month, 0).getDate();

    for (let i = 0; i < 42; i++) {
      const dayOffset = i - startOffset + 1;
      let cellYear = year, cellMonth = month, day = dayOffset, muted = false;
      if (day < 1) { cellMonth--; if (cellMonth < 0) { cellMonth = 11; cellYear--; } day = prevDays + day; muted = true; }
      else if (day > daysInMonth) { day -= daysInMonth; cellMonth++; if (cellMonth > 11) { cellMonth = 0; cellYear++; } muted = true; }

      const key = cellYear + "-" + String(cellMonth + 1).padStart(2, "0") + "-" + String(day).padStart(2, "0");
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = day;
      if (muted) button.classList.add("is-muted");
      const available = availableReportDates.includes(key);
      if (available) button.classList.add("is-available");
      if (key === els.reportDate.value) button.classList.add("is-selected");
      const todayKey = dateKeyFromValue(new Date());
      if (key === todayKey) button.classList.add("is-today");
      button.disabled = !available;
      if (available) {
        button.addEventListener("click", () => {
          els.reportDate.value = key;
          els.datePickerValue.textContent = formatPickerDate(key);
          els.processBtn.disabled = false;
          els.datePickerPopover.classList.add("hidden");
          els.datePickerButton.setAttribute("aria-expanded", "false");
          renderCalendar();
        });
      }
      els.calendarGrid.appendChild(button);
    }
  }

  function shiftCalendar(months) {
    calendarCursor = new Date(calendarCursor.getFullYear(), calendarCursor.getMonth() + months, 1);
    renderCalendar();
  }

  function closeCalendar() {
    els.datePickerPopover.classList.add("hidden");
    els.datePickerButton.setAttribute("aria-expanded", "false");
  }

  function processSelectedDate() {
    const date = clean(els.reportDate.value);
    if (!date) return alert("Please select a report date first.");
    selectedReportDate = date;
    const dateRows = rawRows.filter(row => dateKeyFromValue(sourceValue(row, "Created on")) === date);
    const nabfidRows = dateRows.filter(isNabfidRecord);
    const socAlertRows = nabfidRows.filter(isSocAlert);
    const excludedRuleRows = socAlertRows.filter(isExcludedNabfidRule).length;
    const blankResolutionRows = socAlertRows.filter(row => !isExcludedNabfidRule(row) && isBlankResolution(row)).length;
    const eligibleRows = socAlertRows.filter(row => !isExcludedNabfidRule(row) && !isBlankResolution(row));

    if (!eligibleRows.length) {
      processedRows = [];
      els.resultSection.classList.add("hidden");
      els.dateSelectionStatus.textContent = "No eligible NaBFID records found for the selected date.";
      return;
    }
    processedRows = processSheet(eligibleRows);
    renderSummary();
    els.fileName.textContent = "Processed: " + processedRows.length + " NaBFID records for " + formatEmailDate(processedRows[0]["Created on"]);
    els.dateSelectionStatus.textContent = processedRows.length + " records selected • " + excludedRuleRows + " excluded by rule • " + blankResolutionRows + " excluded with blank resolution steps.";
    els.resultSection.classList.remove("hidden");
    els.dataPanel.classList.add("hidden");
    requestAnimationFrame(() => els.resultSection.scrollIntoView({ behavior: "smooth", block: "start" }));
  }

  async function handleFile(file) {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx", "xls"].includes(ext)) return alert("Please choose an Excel file (.xlsx or .xls).");
    if (typeof XLSX === "undefined") return alert("Excel engine is not loaded. Please refresh the page and try again.");
    resetForNewUpload();
    els.fileName.textContent = "Reading: " + file.name;
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false, cellNF: true, cellText: true });
      if (!workbook.SheetNames?.length) throw new Error("No sheets found in workbook.");
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      if (!sheet) throw new Error("Unable to read the first worksheet.");
      rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
      if (!rawRows.length) throw new Error("The workbook is empty.");
      if (!Object.keys(rawRows[0]).some(h => clean(h).toLowerCase() === "subject")) throw new Error('Required column "subject" was not found.');
      if (!Object.keys(rawRows[0]).some(h => clean(h).toLowerCase() === "created on")) throw new Error('Required column "Created on" was not found.');
      if (!Object.keys(rawRows[0]).some(h => clean(h).toLowerCase() === "resolution_steps")) throw new Error('Required column "resolution_steps" was not found.');
      showDateSelection(rawRows);
      els.fileName.textContent = "Ready: " + file.name + " • " + rawRows.length + " raw records";
      requestAnimationFrame(() => els.dateSelection.scrollIntoView({ behavior: "smooth", block: "center" }));
    } catch (error) {
      resetForNewUpload();
      console.error("Offense Processor:", error);
      els.fileName.textContent = "Processing failed: " + file.name;
      alert("Could not read file: " + (error?.message || error));
    }
  }

  function bindEvents() {
    els.fileInput.addEventListener("change", e => e.target.files?.[0] && handleFile(e.target.files[0]));
    els.datePickerButton.addEventListener("click", () => {
      const open = !els.datePickerPopover.classList.contains("hidden");
      if (open) closeCalendar();
      else {
        els.datePickerPopover.classList.remove("hidden");
        els.datePickerButton.setAttribute("aria-expanded", "true");
        renderCalendar();
      }
    });
    els.calendarPrev.addEventListener("click", () => shiftCalendar(-1));
    els.calendarNext.addEventListener("click", () => shiftCalendar(1));
    els.calendarToday.addEventListener("click", () => {
      const todayKey = dateKeyFromValue(new Date());
      if (availableReportDates.includes(todayKey)) {
        els.reportDate.value = todayKey;
        els.datePickerValue.textContent = formatPickerDate(todayKey);
        els.processBtn.disabled = false;
      }
      calendarCursor = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
      renderCalendar();
    });
    document.addEventListener("click", event => {
      if (!els.datePickerPopover.contains(event.target) && !els.datePickerButton.contains(event.target)) closeCalendar();
    });
    els.processBtn.addEventListener("click", processSelectedDate);
    els.dropzone.addEventListener("dragover", e => {
      e.preventDefault();
      els.dropzone.classList.add("drop-active");
    });
    els.dropzone.addEventListener("dragleave", () => els.dropzone.classList.remove("drop-active"));
    els.dropzone.addEventListener("drop", e => {
      e.preventDefault();
      els.dropzone.classList.remove("drop-active");
      if (e.dataTransfer.files?.[0]) handleFile(e.dataTransfer.files[0]);
    });
    els.previewBtn.addEventListener("click", () => showDataPanel(false));
    els.editBtn.addEventListener("click", () => showDataPanel(true));
    els.copyTableBtn.addEventListener("click", copySummaryTable);
    els.downloadBtn.addEventListener("click", download);
    els.copyEmailBtn.addEventListener("click", copyEmail);
    els.copySubjectBtn.addEventListener("click", copySubject);
    els.saveBtn.addEventListener("click", () => {
      recalculateDerivedFields();
      renderSummary();
      renderData(true);
    });
    els.closePanelBtn.addEventListener("click", () => els.dataPanel.classList.add("hidden"));
  }

  function init() {
    try {
      assertUI();
      bindEvents();
      if (typeof XLSX === "undefined") {
        els.fileName.textContent = "Excel engine unavailable — refresh the page";
        console.error("Offense Processor: XLSX library is not loaded.");
      }
    } catch (error) {
      console.error("Offense Processor initialization failed:", error);
      alert("The application could not initialize. Please refresh the page.");
    }
  }

  init();
})();