"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const expectedTicketIds = [
  "SOCALQR8908445", "SOCALQR8908501", "SOCALQR8909571", "SOCALQR8909725",
  "SOCALQR8909737", "SOCALQR8909757", "SOCALQR8910015", "SOCALQR8910203",
  "SOCALQR8912559", "SOCALQR8913143", "SOCALQR8914061", "SOCALQR8914107",
  "SOCALQR8917403", "SOCALQR8917745", "SOCALQR8918949", "SOCALQR8918959",
  "SOCALQR8919037", "SOCALQR8919719", "SOCALQR8919881", "SOCALQR8920835",
  "SOCALQR8921107", "SOCALQR8921317", "SOCALQR8921399", "SOCALQR8922725",
  "SOCALQR8922943", "SOCALQR8925427", "SOCALQR8926439", "SOCALQR8926445",
  "SOCALQR8926473", "SOCALQR8926771", "SOCALQR8928631", "SOCALQR8931177",
  "SOCALQR8931433", "SOCALQR8931723"
];

const application = fs.readFileSync("app.js", "utf8").replace(
  "  init();\n})();",
  "  globalThis.__offenseTest = { dateKeyFromValue, isNabfidRecord, isExcludedNabfidRule, isBlankResolution, processSheet };\n})();"
);
const context = { document: { getElementById: () => ({}) }, console };
context.globalThis = context;
vm.runInNewContext(application, context, { filename: "app.js" });
const rules = context.__offenseTest;

// This mirrors the object shape emitted by XLSX.utils.sheet_to_json with
// { defval: "", raw: false }.  The two additional rows model duplicate
// headings, for which SheetJS emits subject_1 and resolution_steps_1.
const parsedRows = expectedTicketIds.map(ticket => ({
  "Tickets#": ticket,
  "Created on": "22-09-2026 09:00:00 AM",
  subject: "NaBFID | routine offense",
  Organization: "NaBFID",
  resolution_steps: "Investigated",
}));
parsedRows.push(
  {
    "Tickets#": "excluded-duplicate-subject",
    "Created on": "22-09-2026 10:00:00 AM",
    subject: "NaBFID | routine offense",
    subject_1: "NaBFID devices\u200B not\u200B working",
    Organization: "NaBFID",
    resolution_steps: "Investigated",
  },
  {
    "Tickets#": "excluded-blank-resolution",
    "Created on": "22-09-2026 11:00:00 AM",
    subject: "NaBFID | routine offense",
    Organization: "NaBFID",
    resolution_steps: "\u200B",
    resolution_steps_1: "\t",
  }
);

const eligibleRows = parsedRows.filter(row =>
  rules.dateKeyFromValue(row["Created on"]) === "2026-09-22" &&
  rules.isNabfidRecord(row) &&
  !rules.isExcludedNabfidRule(row) &&
  !rules.isBlankResolution(row)
);

assert.equal(parsedRows.length, 36);
assert.deepEqual([...eligibleRows.map(row => row["Tickets#"])], expectedTicketIds);
assert.equal(rules.processSheet(eligibleRows).length, 34);
console.log("Offense filtering regression: 34 eligible records");
