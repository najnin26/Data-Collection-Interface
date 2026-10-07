/* ===================================================================
   STORAGE + EXCEL EXPORT
   Uses localStorage for persistence + SheetJS for Excel download
   =================================================================== */

const STORAGE_KEYS = {
  RESPONSES: "bp_responses",
  USABILITY: "bp_usability"
};

/* ----------------------- UUID generator --------------------------- */
function generateUUID() {
  if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, c => {
    const r = (Math.random() * 16) | 0;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

/* ----------------------- Sanitize text ---------------------------- */
function safeText(x) {
  return String(x ?? "").replace(/\x00/g, "").trim();
}

/* ----------------------- Local storage helpers -------------------- */
function getResponses() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEYS.RESPONSES)) || [];
  } catch {
    return [];
  }
}

function getUsability() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEYS.USABILITY)) || [];
  } catch {
    return [];
  }
}

function saveResponseToStorage(row) {
  const data = getResponses();
  data.push(row);
  localStorage.setItem(STORAGE_KEYS.RESPONSES, JSON.stringify(data));
}

function saveUsabilityToStorage(row) {
  const data = getUsability();
  data.push(row);
  localStorage.setItem(STORAGE_KEYS.USABILITY, JSON.stringify(data));
}

/* ----------------------- Excel export ----------------------------- */
function downloadResponsesExcel() {
  const data = getResponses();

  if (!data.length) {
    alert("No responses collected yet.");
    return;
  }

  const headers = [
    "Response_ID",
    "Participant_ID",
    "Timestamp",
    "Participant_Age_Group",
    "Education",
    "Bangla_Usage",
    "Bangla_Variety",
    "Scenario_ID",
    "Scenario_Category",
    "Context",
    "Utterance",
    "Intended_Meaning",
    "Primary_Intent_Category",
    "Scenario_Response_Time_Seconds"
  ];

  const rows = data.map(r => [
    r.response_id, r.participant_id, r.timestamp,
    r.age, r.education, r.usage, r.variety,
    r.scenario_id, r.category, r.context,
    r.utterance, r.intended, r.intent_cat,
    r.response_time
  ]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws["!cols"] = [
    { wch: 38 }, { wch: 38 }, { wch: 22 }, { wch: 16 },
    { wch: 18 }, { wch: 18 }, { wch: 30 },
    { wch: 10 }, { wch: 26 }, { wch: 60 },
    { wch: 55 }, { wch: 55 }, { wch: 30 }, { wch: 16 }
  ];
  XLSX.utils.book_append_sheet(wb, ws, "Responses");

  const ts = new Date().toISOString().replace(/[:.]/g, "-");
  XLSX.writeFile(wb, `bangla_pragmatics_responses_${ts}.xlsx`);
}

function downloadUsabilityExcel() {
  const data = getUsability();

  if (!data.length) {
    alert("No usability feedback collected yet.");
    return;
  }

  const headers = [
    "Participant_ID",
    "Completion_Time_Minutes",
    "U1_Instructions_Clear",
    "U2_Scenarios_Understandable",
    "U3_Interface_Easy",
    "U4_Example_Helpful",
    "U5_Progress_Useful",
    "U6_Easy_to_Express_Meaning",
    "U7_Questionnaire_Length_Reasonable",
    "U8_Comfortable_Using_System",
    "U9_Willing_to_Reuse",
    "U10_Overall_Satisfaction",
    "Open_Feedback"
  ];

  const rows = data.map(r => [
    r.participant_id, r.completion_time,
    r.u1, r.u2, r.u3, r.u4, r.u5,
    r.u6, r.u7, r.u8, r.u9, r.u10,
    r.feedback
  ]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws["!cols"] = [
    { wch: 38 }, { wch: 22 },
    { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
    { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
    { wch: 60 }
  ];
  XLSX.utils.book_append_sheet(wb, ws, "Usability_Feedback");

  const ts = new Date().toISOString().replace(/[:.]/g, "-");
  XLSX.writeFile(wb, `bangla_pragmatics_usability_${ts}.xlsx`);
}