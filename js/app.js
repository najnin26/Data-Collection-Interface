/* ===================================================================
   MAIN APP LOGIC
   =================================================================== */

const state = {
  participant: null,       // { participant_id, age, education, usage, variety, start_time }
  index: 0,                // scenario index
  answers: {},             // { [scenarioIdx]: { utterance, intended, cat_idx, saved } }
  scenarioStartTime: null, // Date when current scenario was shown
  usabilityAnswers: {}     // { u1: 5, u2: 4, ... }
};

/* ------------------------ Page routing ---------------------------- */
function showPage(pageId) {
  document.querySelectorAll(".page").forEach(p => p.classList.remove("active"));
  document.getElementById(pageId).classList.add("active");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

/* ------------------------ Initialize ------------------------------ */
document.addEventListener("DOMContentLoaded", () => {
  buildIntentOptions();
  buildUsabilityForm();
  bindParticipantForm();
  bindScenarioForm();
  bindUsabilityForm();
  bindThankYouButtons();
  showPage("page-participant");
});

/* ------------------------ Participant form ------------------------ */
function bindParticipantForm() {
  const form = document.getElementById("participant-form");
  const errorEl = document.getElementById("participant-error");

  form.addEventListener("submit", e => {
    e.preventDefault();
    errorEl.textContent = "";

    const consent = document.getElementById("consent").checked;
    if (!consent) {
      errorEl.textContent = "Please provide consent before starting.";
      return;
    }

    state.participant = {
      participant_id: generateUUID(),
      age: document.getElementById("age").value,
      education: document.getElementById("education").value,
      usage: document.getElementById("usage").value,
      variety: document.getElementById("variety").value,
      start_time: new Date().toISOString()
    };

    state.index = 0;
    state.answers = {};
    state.scenarioStartTime = new Date();

    renderScenario();
    showPage("page-scenario");
  });
}

/* ------------------------ Intent options -------------------------- */
function buildIntentOptions() {
  const sel = document.getElementById("intent-cat");
  sel.innerHTML = "";
  INTENT_OPTIONS.forEach((opt, i) => {
    const o = document.createElement("option");
    o.value = opt;
    o.textContent = opt;
    if (i === 0) o.selected = true;
    sel.appendChild(o);
  });
}

/* ------------------------ Render scenario ------------------------- */
function renderScenario() {
  const i = state.index;
  const total = SCENARIOS.length;
  const scenario = SCENARIOS[i];

  // Progress bar
  const progress = ((i + 1) / total) * 100;
  document.getElementById("progress-bar").style.width = progress + "%";

  // Header
  document.getElementById("scenario-counter").textContent =
    `Scenario ${i + 1} of ${total}`;

  // Break notice
  const breakEl = document.getElementById("break-notice");
  breakEl.style.display = (i === 25) ? "block" : "none";

  // Context
  document.getElementById("scenario-context").textContent = scenario.context;

  // Pre-fill if previously answered
  const ans = state.answers[i] || {};
  document.getElementById("utterance").value = ans.utterance || "";
  document.getElementById("intended").value = ans.intended || "";
  document.getElementById("intent-cat").value =
    ans.intent_cat || INTENT_OPTIONS[0];

  // Previous button disabled on first scenario
  document.getElementById("btn-prev").disabled = (i === 0);

  // Clear error
  document.getElementById("scenario-error").textContent = "";
}

/* ------------------------ Scenario form --------------------------- */
function bindScenarioForm() {
  const form = document.getElementById("scenario-form");
  const errorEl = document.getElementById("scenario-error");

  document.getElementById("btn-prev").addEventListener("click", () => {
    if (state.index === 0) return;
    saveCurrentAnswerLocally();
    state.index = Math.max(0, state.index - 1);
    state.scenarioStartTime = new Date(); // reset timer
    renderScenario();
  });

  form.addEventListener("submit", e => {
    e.preventDefault();
    errorEl.textContent = "";

    const utterance = document.getElementById("utterance").value.trim();
    const intended = document.getElementById("intended").value.trim();
    const intentCat = document.getElementById("intent-cat").value;

    if (!utterance || !intended) {
      errorEl.textContent =
        "অনুগ্রহ করে দুটি প্রশ্নেরই উত্তর দিন। (Please answer both text questions.)";
      return;
    }

    const i = state.index;

    // Save locally
    state.answers[i] = {
      utterance, intended,
      intent_cat: intentCat,
      saved: state.answers[i]?.saved || false
    };

    // Save to storage (only once per scenario)
    if (!state.answers[i].saved) {
      const responseTime = state.scenarioStartTime
        ? (Date.now() - state.scenarioStartTime.getTime()) / 1000
        : 0;

      const row = {
        response_id: generateUUID(),
        participant_id: state.participant.participant_id,
        timestamp: new Date().toISOString().replace(/\.\d+Z$/, ""),
        age: state.participant.age,
        education: state.participant.education,
        usage: state.participant.usage,
        variety: state.participant.variety,
        scenario_id: i + 1,
        category: SCENARIOS[i].category,
        context: SCENARIOS[i].context,
        utterance: safeText(utterance),
        intended: safeText(intended),
        intent_cat: safeText(intentCat),
        response_time: +responseTime.toFixed(2)
      };

      saveResponseToStorage(row);
      state.answers[i].saved = true;
    }

    // Next
    state.index += 1;

    if (state.index < SCENARIOS.length) {
      state.scenarioStartTime = new Date();
      renderScenario();
    } else {
      // Move to usability page
      renderUsability();
      showPage("page-usability");
    }
  });
}

function saveCurrentAnswerLocally() {
  const i = state.index;
  state.answers[i] = {
    utterance: document.getElementById("utterance").value,
    intended: document.getElementById("intended").value,
    intent_cat: document.getElementById("intent-cat").value,
    saved: state.answers[i]?.saved || false
  };
}

/* ------------------------ Usability form -------------------------- */
function buildUsabilityForm() {
  const container = document.getElementById("usability-questions");
  container.innerHTML = "";

  USABILITY_QUESTIONS.forEach(q => {
    const div = document.createElement("div");
    div.className = "usability-question";

    const p = document.createElement("p");
    p.textContent = q.text;
    div.appendChild(p);

    const group = document.createElement("div");
    group.className = "radio-group";

    const scaleLabels = [
      "1 - সম্পূর্ণ অসম্মত",
      "2 - অসম্মত",
      "3 - নিরপেক্ষ",
      "4 - সম্মত",
      "5 - সম্পূর্ণ সম্মত"
    ];

    scaleLabels.forEach((label, idx) => {
      const lbl = document.createElement("label");
      const input = document.createElement("input");
      input.type = "radio";
      input.name = q.id;
      input.value = idx + 1;
      input.required = true;

      lbl.appendChild(input);
      lbl.appendChild(document.createTextNode(label));
      group.appendChild(lbl);
    });

    div.appendChild(group);
    container.appendChild(div);
  });
}

function renderUsability() {
  // Optionally, nothing needed — form is static
  document.getElementById("feedback").value = "";
}

function bindUsabilityForm() {
  const form = document.getElementById("usability-form");

  form.addEventListener("submit", e => {
    e.preventDefault();

    const answers = {};
    let allAnswered = true;

    USABILITY_QUESTIONS.forEach(q => {
      const sel = document.querySelector(`input[name="${q.id}"]:checked`);
      if (!sel) allAnswered = false;
      answers[q.id] = sel ? parseInt(sel.value, 10) : null;
    });

    if (!allAnswered) {
      alert("Please answer all 10 usability questions.");
      return;
    }

    const feedback = document.getElementById("feedback").value;

    // Completion time (minutes)
    const completionTime =
      (Date.now() - new Date(state.participant.start_time).getTime()) /
      1000 / 60;

    const row = {
      participant_id: state.participant.participant_id,
      completion_time: +completionTime.toFixed(2),
      u1: answers.u1, u2: answers.u2, u3: answers.u3, u4: answers.u4,
      u5: answers.u5, u6: answers.u6, u7: answers.u7, u8: answers.u8,
      u9: answers.u9, u10: answers.u10,
      feedback: safeText(feedback)
    };

    saveUsabilityToStorage(row);

    // Confetti-ish
    if (typeof confetti === "function") confetti();

    showPage("page-thanks");
  });
}

/* ------------------------ Thank you page -------------------------- */
function bindThankYouButtons() {
  document.getElementById("btn-download")
    .addEventListener("click", downloadResponsesExcel);

  document.getElementById("btn-download-usability")
    .addEventListener("click", downloadUsabilityExcel);

  document.getElementById("btn-new-response").addEventListener("click", () => {
    // Reset state
    state.participant = null;
    state.index = 0;
    state.answers = {};
    state.scenarioStartTime = null;

    // Reset forms
    document.getElementById("participant-form").reset();
    document.getElementById("scenario-form").reset();
    document.getElementById("usability-form").reset();

    showPage("page-participant");
  });
}