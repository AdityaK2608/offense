(function(){
  const ORGANIZATIONS = {
    kpmg: "KPMG Assurance and Consulting Services LLP",
    default: "Edgeverve Systems Limited- NaBFID"
  };
  let processedRows = [];
  const EXCEL_FONT = "Inter 18pt";

  const els = {
    fileInput: document.getElementById("fileInput"),
    browseBtn: document.getElementById("browseBtn"),
    dropzone: document.getElementById("dropzone"),
    fileName: document.getElementById("fileName"),
    resultSection: document.getElementById("resultSection"),
    previewBtn: document.getElementById("previewBtn"),
    editBtn: document.getElementById("editBtn"),
    downloadBtn: document.getElementById("downloadBtn"),
    copyTableBtn: document.getElementById("copyTableBtn"),
    dataPanel: document.getElementById("dataPanel"),
    dataTitle: document.getElementById("dataTitle"),
    dataSubtitle: document.getElementById("dataSubtitle"),
    dataHead: document.getElementById("dataHead"),
    dataBody: document.getElementById("dataBody"),
    saveBtn: document.getElementById("saveBtn"),
    closePanelBtn: document.getElementById("closePanelBtn")
  };

  function clean(v){ return v == null ? "" : String(v).trim(); }

  function extractClassification(subject){
    const match = clean(subject).match(/Domain:\s*([^|]*)/i);
    const domain = match ? clean(match[1]) : "";
    return domain || "NABFID DC";
  }

  function getOrganization(subject){
    return /KPMG/i.test(clean(subject))
      ? ORGANIZATIONS.kpmg
      : ORGANIZATIONS.default;
  }

  function processSheet(rawRows){
    if (!rawRows.length) throw new Error("The workbook is empty.");

    const originalHeaders = Object.keys(rawRows[0]);
    const subjectKey = originalHeaders.find(
      h => h.trim().toLowerCase() === "subject"
    );

    if (!subjectKey) throw new Error('Required column "subject" was not found.');

    return rawRows.map(row => {
      const next = {};
      originalHeaders.forEach(h => {
        next[h] = clean(row[h]);
        if (h === "subject") {
          next["Classification"] = extractClassification(row[subjectKey]);
          next["Organization"] = getOrganization(row[subjectKey]);
        }
      });
      return reorderColumns(next, originalHeaders);
    });
  }

  function reorderColumns(row, originalHeaders){
    const columns = [];
    originalHeaders.forEach(h => {
      columns.push(h);
      if (h === "subject") columns.push("Classification", "Organization");
    });

    if (!columns.includes("Classification")) columns.push("Classification");
    if (!columns.includes("Organization")) columns.push("Organization");

    const unique = [...new Set(columns)].filter(
      c => Object.prototype.hasOwnProperty.call(row, c)
    );

    const result = {};
    unique.forEach(c => result[c] = row[c]);
    return result;
  }

  function normalizeStatus(value){
    return clean(value).toLowerCase();
  }

  function summarize(rows){
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

  function renderSummary(){
    const map = summarize(processedRows);
    const thead = document.getElementById("summaryHead");
    const tbody = document.getElementById("summaryBody");

    const headers = [
      "Client",
      "Closed",
      "Pending on COE",
      "Pending On Customer",
      "Grand Total"
    ];

    thead.innerHTML = "<tr>" + headers.map(
      h => '<th class="bg-slate-900 px-4 py-3 text-left font-semibold text-white">' + h + "</th>"
    ).join("") + "</tr>";

    const totals = {
      Closed: 0,
      "Pending on COE": 0,
      "Pending On Customer": 0,
      total: 0
    };

    tbody.innerHTML = "";

    Array.from(map.keys()).sort().forEach(client => {
      const summary = map.get(client);

      totals.Closed += summary.Closed;
      totals["Pending on COE"] += summary["Pending on COE"];
      totals["Pending On Customer"] += summary["Pending On Customer"];
      totals.total += summary.total;

      const cells = [
        client,
        summary.Closed,
        summary["Pending on COE"],
        summary["Pending On Customer"],
        summary.total
      ];

      const tr = document.createElement("tr");
      tr.className = "border-b border-slate-200";

      cells.forEach((value, index) => {
        const td = document.createElement("td");
        td.className = "px-4 py-3" + (index === 0 ? " font-medium" : "");
        td.textContent = value;
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    const totalRow = document.createElement("tr");
    totalRow.className = "bg-slate-900 font-semibold text-white";

    [
      "Grand Total",
      totals.Closed,
      totals["Pending on COE"],
      totals["Pending On Customer"],
      totals.total
    ].forEach(value => {
      const td = document.createElement("td");
      td.className = "px-4 py-3";
      td.textContent = value;
      totalRow.appendChild(td);
    });

    tbody.appendChild(totalRow);

    document.getElementById("statRecords").textContent = processedRows.length;
    document.getElementById("statClients").textContent = map.size;
    document.getElementById("statClosed").textContent = totals.Closed;
    document.getElementById("statPending").textContent =
      totals["Pending on COE"] + totals["Pending On Customer"];
  }

  function renderData(editable){
    if (!processedRows.length) return;

    els.dataPanel.classList.remove("hidden");
    els.saveBtn.classList.toggle("hidden", !editable);
    els.dataTitle.textContent = editable ? "Edit processed data" : "Preview processed data";
    els.dataSubtitle.textContent = editable
      ? "Edit cells below, then save edits to refresh the summary."
      : "Read-only preview of the generated Offenses sheet.";

    const headers = Object.keys(processedRows[0]);

    els.dataHead.innerHTML = "<tr>" + headers.map(
      h => '<th class="border-b border-slate-200 px-4 py-3 text-left font-semibold">' + h + "</th>"
    ).join("") + "</tr>";

    els.dataBody.innerHTML = "";

    processedRows.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");
      tr.className = "border-b border-slate-100";

      headers.forEach(header => {
        const td = document.createElement("td");
        td.className = "data-cell px-4 py-2 align-top";

        if (editable) {
          const input = document.createElement("input");
          input.value = row[header] || "";
          input.className = "w-full rounded-lg border border-slate-300 px-3 py-2 outline-none focus:border-slate-500 focus:ring-2 focus:ring-slate-200";

          input.addEventListener("input", event => {
            processedRows[rowIndex][header] = event.target.value;
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

  async function handleFile(file){
    if (!file) return;

    if (!/\.xlsx?$/.test(file.name.toLowerCase())) {
      alert("Please choose an Excel file (.xlsx or .xls).");
      return;
    }

    els.fileName.textContent = "Selected: " + file.name;

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: false, cellNF: true, cellText: true });

      if (!workbook.SheetNames.length) {
        throw new Error("No sheets found in workbook.");
      }

      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: false });

      processedRows = processSheet(rawRows);
      renderSummary();

      els.resultSection.classList.remove("hidden");
      els.dataPanel.classList.add("hidden");

      window.scrollTo({
        top: document.body.scrollHeight,
        behavior: "smooth"
      });
    } catch (error) {
      alert("Could not process file: " + error.message);
      console.error(error);
    }
  }

  function download(){
    if (!processedRows.length) return;

    const workbook = XLSX.utils.book_new();

    const border = {
      top: { style: "thin", color: { rgb: "D9DEE7" } },
      bottom: { style: "thin", color: { rgb: "D9DEE7" } },
      left: { style: "thin", color: { rgb: "D9DEE7" } },
      right: { style: "thin", color: { rgb: "D9DEE7" } }
    };

    const headerStyle = {
      font: { name: EXCEL_FONT, sz: 10, family: 2, bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "0B2A5B" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: false },
      border: {
        top: { style: "thin", color: { rgb: "202020" } },
        bottom: { style: "thin", color: { rgb: "202020" } },
        left: { style: "thin", color: { rgb: "202020" } },
        right: { style: "thin", color: { rgb: "202020" } }
      }
    };

    const bodyStyle = {
      font: { name: EXCEL_FONT, sz: 10, family: 2, color: { rgb: "374151" } },
      fill: { fgColor: { rgb: "FFFFFF" } },
      alignment: { vertical: "center", wrapText: false },
      border: {
        top: { style: "thin", color: { rgb: "202020" } },
        bottom: { style: "thin", color: { rgb: "202020" } },
        left: { style: "thin", color: { rgb: "202020" } },
        right: { style: "thin", color: { rgb: "202020" } }
      }
    };

    const alternateStyle = {
      font: { name: EXCEL_FONT, sz: 10, family: 2, color: { rgb: "374151" } },
      fill: { fgColor: { rgb: "FFFFFF" } },
      alignment: { vertical: "center", wrapText: false },
      border: {
        top: { style: "thin", color: { rgb: "202020" } },
        bottom: { style: "thin", color: { rgb: "202020" } },
        left: { style: "thin", color: { rgb: "202020" } },
        right: { style: "thin", color: { rgb: "202020" } }
      }
    };

    // -----------------------------
    // Offenses sheet
    // -----------------------------
    const offenseSheet = XLSX.utils.json_to_sheet(processedRows);
    offenseSheet["!freeze"] = { xSplit: 0, ySplit: 1 };
    offenseSheet["!autofilter"] = { ref: offenseSheet["!ref"] };

    const offenseRange = XLSX.utils.decode_range(offenseSheet["!ref"]);
    const offenseHeaders = Object.keys(processedRows[0]);

    // Sensible widths: readable without creating an enormous worksheet.
    const widthByHeader = {
      "Tickets#": 16,
      "Created on": 20,
      "Department": 20,
      "Prioritytitle": 16,
      "Type": 16,
      "subject": 42,
      "Classification": 22,
      "Organization": 38,
      "Wing": 20,
      "Closedon": 20,
      "resolution_steps": 55,
      "Status": 18
    };

    offenseSheet["!cols"] = offenseHeaders.map(header => ({
      wch: widthByHeader[header] || Math.min(
        36,
        Math.max(
          14,
          header.length + 2,
          ...processedRows.slice(0, 100).map(row => clean(row[header]).length + 2)
        )
      )
    }));

    for (let row = offenseRange.s.r; row <= offenseRange.e.r; row++) {
      offenseSheet["!rows"] = offenseSheet["!rows"] || [];
      offenseSheet["!rows"][row] = { hpt: row === 0 ? 24 : 20 };

      for (let col = offenseRange.s.c; col <= offenseRange.e.c; col++) {
        const address = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = offenseSheet[address];
        if (!cell) continue;

        if (row === 0) {
          cell.s = headerStyle;
        } else {
          cell.s = bodyStyle;
        }
      }
    }

    XLSX.utils.book_append_sheet(workbook, offenseSheet, "Offenses");

    // -----------------------------
    // Classification summary sheet
    // -----------------------------
    const map = summarize(processedRows);
    const summaryRows = [[
      "Client",
      "Closed",
      "Pending on COE",
      "Pending On Customer",
      "Grand Total"
    ]];

    const grandTotals = [0, 0, 0, 0];

    Array.from(map.keys()).sort().forEach(client => {
      const summary = map.get(client);

      summaryRows.push([
        client,
        summary.Closed,
        summary["Pending on COE"],
        summary["Pending On Customer"],
        summary.total
      ]);

      grandTotals[0] += summary.Closed;
      grandTotals[1] += summary["Pending on COE"];
      grandTotals[2] += summary["Pending On Customer"];
      grandTotals[3] += summary.total;
    });

    summaryRows.push([
      "Grand Total",
      grandTotals[0],
      grandTotals[1],
      grandTotals[2],
      grandTotals[3]
    ]);

    const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
    summarySheet["!freeze"] = { xSplit: 0, ySplit: 1 };
    summarySheet["!autofilter"] = {
      ref: "A1:E" + summaryRows.length
    };
    summarySheet["!cols"] = [
      { wch: 30 },
      { wch: 14 },
      { wch: 20 },
      { wch: 24 },
      { wch: 16 }
    ];

    const summaryRange = XLSX.utils.decode_range(summarySheet["!ref"]);

    const totalStyle = {
      font: { name: EXCEL_FONT, sz: 10, family: 2, bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "0B2A5B" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: false },
      border: {
        top: { style: "thin", color: { rgb: "202020" } },
        bottom: { style: "thin", color: { rgb: "202020" } },
        left: { style: "thin", color: { rgb: "202020" } },
        right: { style: "thin", color: { rgb: "202020" } }
      }
    };

    for (let row = summaryRange.s.r; row <= summaryRange.e.r; row++) {
      summarySheet["!rows"] = summarySheet["!rows"] || [];
      summarySheet["!rows"][row] = { hpt: row === 0 ? 24 : 22 };

      for (let col = summaryRange.s.c; col <= summaryRange.e.c; col++) {
        const address = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = summarySheet[address];
        if (!cell) continue;

        if (row === 0) {
          cell.s = headerStyle;
        } else if (row === summaryRange.e.r) {
          cell.s = totalStyle;
        } else {
          cell.s = {
            ...bodyStyle,
            fill: { fgColor: { rgb: "FFFFFF" } },
            alignment: {
              horizontal: col === 0 ? "left" : "center",
              vertical: "center",
              wrapText: false
            },
            border: {
              top: { style: "thin", color: { rgb: "202020" } },
              bottom: { style: "thin", color: { rgb: "202020" } },
              left: { style: "thin", color: { rgb: "202020" } },
              right: { style: "thin", color: { rgb: "202020" } }
            }
          };
        }
      }
    }

    XLSX.utils.book_append_sheet(workbook, summarySheet, "Classification");

    const output = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array"
    });

    const blob = new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);

    function formatFileDate(value){
      const text = clean(value);
      if (!text) return "";

      // Prefer the date portion from the Created on value.
      const match = text.match(/(\\d{1,4})[\\/\\-.](\\d{1,2})[\\/\\-.](\\d{1,4})/);
      if (!match) return "";

      let first = Number(match[1]);
      let second = Number(match[2]);
      let third = Number(match[3]);
      let day, month, year;

      if (first >= 1000) {
        year = first;
        month = second;
        day = third;
      } else if (third >= 1000) {
        day = first;
        month = second;
        year = third;
      } else {
        return "";
      }

      if (day < 1 || day > 31 || month < 1 || month > 12) return "";

      return String(day).padStart(2, "0") +
        String(month).padStart(2, "0") +
        String(year).slice(-4);
    }

    const createdDate = formatFileDate(processedRows[0]["Created on"]);
    link.href = url;
    link.download = createdDate
      ? "NABFID Offenses - " + createdDate + ".xlsx"
      : "NABFID Offenses.xlsx";
    link.click();

    URL.revokeObjectURL(url);
  }

  els.browseBtn.addEventListener("click", () => els.fileInput.click());

  els.dropzone.addEventListener("click", event => {
    if (event.target !== els.browseBtn) els.fileInput.click();
  });

  els.fileInput.addEventListener("change", event => {
    handleFile(event.target.files[0]);
  });

  ["dragenter", "dragover"].forEach(eventName => {
    els.dropzone.addEventListener(eventName, event => {
      event.preventDefault();
      els.dropzone.classList.add("drop-active");
    });
  });

  ["dragleave", "drop"].forEach(eventName => {
    els.dropzone.addEventListener(eventName, event => {
      event.preventDefault();
      els.dropzone.classList.remove("drop-active");
    });
  });

  els.dropzone.addEventListener("drop", event => {
    handleFile(event.dataTransfer.files[0]);
  });

  async function copySummaryTable(){
    const table = document.querySelector("#summaryHead").closest("table");
    if (!table) return;

    const rows = Array.from(table.querySelectorAll("tr"));
    const bodyRows = rows.slice(1, -1);
    const totalRow = rows[rows.length - 1];

    const border = "#202020";
    const headerBg = "#0B2A5B";
    const totalBg = "#0B2A5B";
    const textColor = "#1F2937";
    const white = "#FFFFFF";
    const font = "Inter 18pt";

    const cellHtml = (cell, isHeader = false, isTotal = false) => {
      const bg = isHeader ? headerBg : isTotal ? totalBg : white;
      const color = isHeader || isTotal ? white : textColor;
      const weight = isHeader || isTotal ? "700" : "400";

      return '<td valign="middle" bgcolor="' + bg + '" style="background-color:' + bg + ' !important;color:' + color + ' !important;font-family:"' + font + '",Inter,Arial,sans-serif !important;font-size:10pt !important;font-weight:' + weight + ' !important;border:1px solid ' + border + ' !important;border-top:1px solid ' + border + ' !important;border-right:1px solid ' + border + ' !important;border-bottom:1px solid ' + border + ' !important;border-left:1px solid ' + border + ' !important;mso-border-alt:solid ' + border + ' 1px !important;padding:12px 16px !important;line-height:1.4 !important;text-align:center !important;white-space:nowrap !important;mso-font-alt:Arial;mso-padding-alt:12px 16px;">' +
        '<div align="center" style="text-align:center !important;width:100%;">' +
        '<font face="' + font + '" color="' + color + '" size="2" style="font-family:"' + font + '",Inter,Arial,sans-serif;color:' + color + ';font-size:10pt;font-weight:' + weight + ';text-align:center;">' +
        clean(cell.textContent) +
        '</font></div></td>';
    };

    let html = '<table width="100%" border="1" bordercolor="#202020" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border-spacing:0;border:1px solid #202020;mso-table-lspace:0pt;mso-table-rspace:0pt;font-family:"' + font + '",Inter,Arial,sans-serif;font-size:10pt;table-layout:fixed;">' +
      '<colgroup><col width="22%"><col width="13%"><col width="20%"><col width="29%"><col width="16%"></colgroup>';

    html += '<thead><tr>';
    rows[0].querySelectorAll("th,td").forEach(cell => {
      html += cellHtml(cell, true, false);
    });
    html += '</tr></thead><tbody>';

    bodyRows.forEach(row => {
      html += '<tr>';
      row.querySelectorAll("th,td").forEach(cell => {
        html += cellHtml(cell, false, false);
      });
      html += '</tr>';
    });

    html += '<tr>';
    totalRow.querySelectorAll("th,td").forEach(cell => {
      html += cellHtml(cell, false, true);
    });
    html += '</tr></tbody></table>';

    const plainRows = rows.map(row =>
      Array.from(row.querySelectorAll("th, td")).map(cell => clean(cell.textContent)).join("\t")
    );
    const text = plainRows.join("\n");

    try {
      if (navigator.clipboard && window.ClipboardItem) {
        await navigator.clipboard.write([
          new ClipboardItem({
            "text/html": new Blob([html], { type: "text/html" }),
            "text/plain": new Blob([text], { type: "text/plain" })
          })
        ]);
      } else {
        const container = document.createElement("div");
        container.innerHTML = html;
        container.style.position = "fixed";
        container.style.left = "-9999px";
        container.style.opacity = "0";
        document.body.appendChild(container);

        const range = document.createRange();
        range.selectNodeContents(container);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);
        document.execCommand("copy");
        selection.removeAllRanges();
        container.remove();
      }

      const original = els.copyTableBtn.textContent;
      els.copyTableBtn.textContent = "Copied ✓";
      setTimeout(() => { els.copyTableBtn.textContent = original; }, 1600);
    } catch (error) {
      console.error("Table copy failed:", error);
      alert("Unable to copy the table. Please try again.");
    }
  }
  els.copyTableBtn.addEventListener("click", copySummaryTable);
  function showDataPanel(editable) {
    if (!processedRows.length) {
      alert("Please upload and process an Excel file first.");
      return;
    }

    renderData(editable);
    els.dataPanel.classList.remove("hidden");
    requestAnimationFrame(() => {
      els.dataPanel.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }

  els.previewBtn.onclick = function () {
    showDataPanel(false);
  };

  els.editBtn.onclick = function () {
    showDataPanel(true);
  };

  els.saveBtn.addEventListener("click", () => {
    renderSummary();
    renderData(true);
  });

  els.closePanelBtn.addEventListener("click", () => {
    els.dataPanel.classList.add("hidden");
  });

  els.downloadBtn.addEventListener("click", download);
})();