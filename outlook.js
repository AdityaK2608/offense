(function () {
  "use strict";

  function clean(value) {
    return value == null ? "" : String(value).trim();
  }

  function getSummaryTable() {
    const head = document.getElementById("summaryHead");
    return head ? head.closest("table") : null;
  }

  function buildOutlookBody() {
    const table = getSummaryTable();
    if (!table) return "";

    const rows = Array.from(table.querySelectorAll("tr")).map(row =>
      Array.from(row.querySelectorAll("th, td")).map(cell => clean(cell.textContent))
    );

    if (!rows.length) return "";

    const widths = rows[0].map((_, columnIndex) =>
      Math.max(...rows.map(row => (row[columnIndex] || "").length))
    );

    const tableText = rows.map(row =>
      row.map((value, index) => String(value).padEnd(widths[index])).join("   ")
    ).join("\r\n");

    return (
      "Hi Team,\r\n\r\n" +
      "Please find below the processed offense summary:\r\n\r\n" +
      tableText +
      "\r\n\r\nRegards,"
    );
  }

  function openOutlook() {
    const button = document.getElementById("outlookBtn");
    if (!button) return;

    const table = getSummaryTable();
    if (!table || !table.querySelector("tbody tr")) return;

    const url = new URL("https://outlook.office.com/mail/deeplink/compose");
    url.searchParams.set("subject", "Offense Summary Report");
    url.searchParams.set("body", buildOutlookBody());

    window.open(url.toString(), "_blank", "noopener,noreferrer");
  }

  const outlookButton = document.getElementById("outlookBtn");
  if (outlookButton) {
    outlookButton.addEventListener("click", openOutlook);
  }
})();