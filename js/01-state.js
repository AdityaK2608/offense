/* Shared state and constants. */
const ORG={kpmg:"KPMG Assurance and Consulting Services LLP",default:"Edgeverve Systems Limited- NaBFID"};
const OUT=["Tickets#","Created on","Department","Prioritytitle","Type","subject","Classification","Organization","Wing","Closedon","resolution_steps","Status"];
const SUMMARY=["Client","Closed","Pending on COE","Pending On Customer","Grand Total"];
const EXCLUDED=[/not\s+reporting/i,/devices\s+not\s+working/i,/devices\s+not\s+reporting/i];
const $=id=>document.getElementById(id);
const REQUIRED_HEADERS=["Tickets#","Created on","subject","Type"];
const REVIEW_UPDATE_FIELDS=["Status","Wing","Closedon","resolution_steps"];
const DERIVED_FIELDS=new Set(["Type","Classification","Organization"]);
let createdRows=[],closedRows=[],processedRows=[],reviewRows=[],availableDates=[],selectedDate="";
const el={fileInput:$("fileInput"),dropzone:$("dropzone"),fileName:$("fileName"),dateSelection:$("dateSelection"),reportDate:$("reportDate"),datePickerButton:$("datePickerButton"),datePickerValue:$("datePickerValue"),datePickerPopover:$("datePickerPopover"),calendarHint:$("calendarHint"),calendarGrid:$("calendarGrid"),processBtn:$("processBtn"),dateSelectionStatus:$("dateSelectionStatus"),reviewPanel:$("reviewPanel"),reviewBody:$("reviewBody"),reviewCount:$("reviewCount"),continueBtn:$("continueBtn"),cancelReviewBtn:$("cancelReviewBtn"),resultSection:$("resultSection"),previewBtn:$("previewBtn"),editBtn:$("editBtn"),copyTableBtn:$("copyTableBtn"),downloadBtn:$("downloadBtn"),copyEmailBtn:$("copyEmailBtn"),copySubjectBtn:$("copySubjectBtn"),emailPreview:$("emailPreview"),dataPanel:$("dataPanel"),dataTitle:$("dataTitle"),dataSubtitle:$("dataSubtitle"),dataHead:$("dataHead"),dataBody:$("dataBody"),saveBtn:$("saveBtn"),closePanelBtn:$("closePanelBtn")};
