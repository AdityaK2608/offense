(function () {
  "use strict";

  const SUBJECT = "Offense Summary Report";

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
    const table = getSummaryTable();
    if (!table || !table.querySelector("tbody tr")) return;

    const body = buildOutlookBody();

    // Try the desktop Outlook protocol without navigating the current page.
    const outlookUrl =
      "ms-outlook:compose?subject=" +
      encodeURIComponent(SUBJECT) +
      "&body=" +
      encodeURIComponent(body);

    const link = document.createElement("a");
    link.href = outlookUrl;
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();

    // Give Windows a moment to handle the protocol. If it does not,
    // offer the Outlook web compose as a fallback instead of silently doing nothing.
    window.setTimeout(function () {
      const fallback = window.confirm(
        "Desktop Outlook could not be opened automatically.\n\nOpen Outlook on the web instead?"
      );

      if (fallback) {
        const webUrl = new URL("https://outlook.office.com/mail/deeplink/compose");
        webUrl.searchParams.set("subject", SUBJECT);
        webUrl.searchParams.set("body", body);
        window.open(webUrl.toString(), "_blank", "noopener,noreferrer");
      }
    }, 1200);
  }

  const outlookButton = document.getElementById("outlookBtn");

  if (outlookButton) {
    outlookButton.addEventListener("click", openOutlook);
  }
})();