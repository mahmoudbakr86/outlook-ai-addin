/* global Office, document, localStorage, fetch */

let currentItem = null;
let isComposeMode = false;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  document.getElementById("app").classList.remove("hidden");
  currentItem = Office.context.mailbox.item;

  // Compose forms have a "to" recipients field; read forms don't.
  isComposeMode = currentItem.itemType === Office.MailboxEnums.ItemType.Message &&
                  typeof currentItem.body.setAsync === "function" &&
                  typeof currentItem.to !== "undefined";

  restoreSettings();
  wireTabs();
  wireDraftTab();
  wireEditTab();
  wireSearchTab();
  wireCategorizeTab();

  document.getElementById("instructions").placeholder = isComposeMode
    ? "e.g. Write a friendly email introducing our new product to a potential client."
    : "e.g. Write a polite reply declining the meeting and proposing next Tuesday instead.";

  document.getElementById("generateBtn").textContent = isComposeMode
    ? "Generate draft"
    : "Generate reply";

  if (!isComposeMode) {
    document.getElementById("insertEditBtn").textContent = "Open reply with edited text";
  }
});

/* ---------- shared helpers ---------- */

function restoreSettings() {
  const apiKeyInput = document.getElementById("apiKey");
  const modelSelect = document.getElementById("model");
  const savedKey = localStorage.getItem("openai_api_key");
  const savedModel = localStorage.getItem("openai_model");
  if (savedKey) apiKeyInput.value = savedKey;
  if (savedModel) modelSelect.value = savedModel;
  document.getElementById("saveKeyBtn").addEventListener("click", saveKey);
}

function saveKey() {
  const key = document.getElementById("apiKey").value.trim();
  const model = document.getElementById("model").value;
  if (!key) {
    setStatus("Enter an API key first.", "error");
    return;
  }
  localStorage.setItem("openai_api_key", key);
  localStorage.setItem("openai_model", model);
  setStatus("Key saved on this device.", "success");
}

function getApiKey() {
  return document.getElementById("apiKey").value.trim() || localStorage.getItem("openai_api_key");
}

function getModel() {
  return document.getElementById("model").value;
}

function setStatus(message, type) {
  const el = document.getElementById("status");
  el.textContent = message || "";
  el.className = "status" + (type ? " " + type : "");
}

function getEmailBodyText() {
  return new Promise((resolve, reject) => {
    currentItem.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

// Generic call to the OpenAI chat completions API.
async function callOpenAI(systemPrompt, userPrompt) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error("Add and save your OpenAI API key at the top first.");

  const response = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: getModel(),
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt }
      ],
      temperature: 0.6
    })
  });

  if (!response.ok) {
    const errBody = await response.json().catch(() => ({}));
    throw new Error(errBody.error?.message || `OpenAI request failed (${response.status})`);
  }

  const data = await response.json();
  return data.choices?.[0]?.message?.content?.trim() || "";
}

function insertWholeBody(text) {
  const htmlBody = text.replace(/\n/g, "<br>");
  return new Promise((resolve, reject) => {
    if (isComposeMode) {
      currentItem.body.setAsync(
        htmlBody,
        { coercionType: Office.CoercionType.Html },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(result.error);
        }
      );
    } else {
      currentItem.displayReplyForm({ htmlBody });
      resolve();
    }
  });
}

/* ---------- tabs ---------- */

function wireTabs() {
  const buttons = document.querySelectorAll(".tab-btn");
  buttons.forEach((btn) => {
    btn.addEventListener("click", () => {
      buttons.forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.add("hidden"));
      document.getElementById("tab-" + btn.dataset.tab).classList.remove("hidden");
      setStatus("");
    });
  });
}

/* ---------- Draft / Reply tab ---------- */

function wireDraftTab() {
  document.getElementById("generateBtn").addEventListener("click", () => generateDraft(false));
  document.getElementById("regenerateBtn").addEventListener("click", () => generateDraft(true));
  document.getElementById("insertBtn").addEventListener("click", insertDraft);
}

async function generateDraft(isRegenerate) {
  const instructions = document.getElementById("instructions").value.trim();
  const tone = document.getElementById("tone").value;
  const outputEl = document.getElementById("output");

  if (!instructions && !isComposeMode) {
    setStatus("Describe what the reply should say.", "error");
    return;
  }

  setStatus("Talking to ChatGPT...", "");
  document.getElementById("generateBtn").disabled = true;
  document.getElementById("regenerateBtn").disabled = true;

  try {
    let systemPrompt;
    let userPrompt;

    if (isComposeMode) {
      systemPrompt = `You write ${tone} emails. Output ONLY the email body text, no subject line, no explanations, no markdown.`;
      userPrompt = instructions || "Write a short, polite email.";
    } else {
      const subject = currentItem.subject || "(no subject)";
      const originalBody = await getEmailBodyText();
      systemPrompt = `You write ${tone} email replies. Output ONLY the reply body text, no subject line, no explanations, no markdown, no unfilled placeholders — write it ready to send.`;
      userPrompt =
        `Original email subject: ${subject}\n\n` +
        `Original email body:\n${originalBody.slice(0, 6000)}\n\n` +
        `Instructions for the reply: ${instructions || "Write an appropriate, helpful reply."}`;
    }

    const text = await callOpenAI(systemPrompt, userPrompt);
    outputEl.value = text;
    document.getElementById("insertBtn").disabled = !text;
    document.getElementById("regenerateBtn").disabled = false;
    setStatus(isRegenerate ? "Regenerated." : "Draft ready. Review, then insert.", "success");
  } catch (err) {
    setStatus("Error: " + err.message, "error");
  } finally {
    document.getElementById("generateBtn").disabled = false;
  }
}

function insertDraft() {
  const text = document.getElementById("output").value;
  if (!text) return;
  const htmlBody = text.replace(/\n/g, "<br>");

  if (isComposeMode) {
    currentItem.body.setSelectedDataAsync(
      htmlBody,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) setStatus("Inserted into email.", "success");
        else setStatus("Could not insert: " + result.error.message, "error");
      }
    );
  } else {
    currentItem.displayReplyForm({ htmlBody });
    setStatus("Opened a reply window with your draft.", "success");
  }
}

/* ---------- Edit tab ---------- */

const EDIT_PROMPTS = {
  grammar: "Fix all grammar, spelling, and punctuation mistakes. Keep the meaning, tone, and length essentially the same.",
  formal: "Rewrite this in a more formal, professional tone. Keep the same meaning and length roughly the same.",
  casual: "Rewrite this in a more casual, friendly tone. Keep the same meaning.",
  shorten: "Make this significantly more concise while keeping the key points and a polite, complete email structure.",
  expand: "Expand this with more detail and a warmer, more complete tone, while staying on topic.",
  summarize: "Summarize this email content into a few short bullet points capturing the key points and any action items."
};

function wireEditTab() {
  document.getElementById("loadEditorBtn").addEventListener("click", loadEditorText);

  document.querySelectorAll(".chip").forEach((chip) => {
    chip.addEventListener("click", () => runEdit(EDIT_PROMPTS[chip.dataset.action]));
  });

  document.getElementById("customEditBtn").addEventListener("click", () => {
    const instruction = document.getElementById("customEdit").value.trim();
    if (!instruction) {
      setStatus("Describe the edit you want.", "error");
      return;
    }
    runEdit(instruction);
  });

  document.getElementById("insertEditBtn").addEventListener("click", async () => {
    const text = document.getElementById("editor").value;
    if (!text) return;
    try {
      await insertWholeBody(text);
      setStatus(isComposeMode ? "Email updated." : "Opened a reply window with the edited text.", "success");
    } catch (err) {
      setStatus("Could not insert: " + err.message, "error");
    }
  });
}

async function loadEditorText() {
  try {
    setStatus("Loading current email text...", "");
    const text = await getEmailBodyText();
    document.getElementById("editor").value = text;
    document.getElementById("insertEditBtn").disabled = false;
    setStatus("Loaded.", "success");
  } catch (err) {
    setStatus("Could not load email text: " + err.message, "error");
  }
}

async function runEdit(instruction) {
  const editorEl = document.getElementById("editor");
  const text = editorEl.value.trim();
  if (!text) {
    setStatus("Load or paste some email text first.", "error");
    return;
  }
  setStatus("Editing with ChatGPT...", "");
  try {
    const result = await callOpenAI(
      "You are an expert email editor. Output ONLY the revised email text, no explanations, no markdown formatting, no preamble.",
      `Instruction: ${instruction}\n\nEmail text:\n${text}`
    );
    editorEl.value = result;
    document.getElementById("insertEditBtn").disabled = !result;
    setStatus("Edit applied.", "success");
  } catch (err) {
    setStatus("Error: " + err.message, "error");
  }
}

/* ---------- Search tab (within the currently open email) ---------- */

function wireSearchTab() {
  document.getElementById("searchBtn").addEventListener("click", runSearch);
  document.getElementById("searchTerm").addEventListener("keydown", (e) => {
    if (e.key === "Enter") runSearch();
  });
}

async function runSearch() {
  const term = document.getElementById("searchTerm").value.trim();
  const resultsEl = document.getElementById("searchResults");
  resultsEl.innerHTML = "";

  if (!term) {
    setStatus("Type a word or phrase to search for.", "error");
    return;
  }

  setStatus("Searching this email...", "");
  try {
    const body = await getEmailBodyText();
    const lower = body.toLowerCase();
    const termLower = term.toLowerCase();

    let matches = [];
    let idx = lower.indexOf(termLower);
    while (idx !== -1) {
      matches.push(idx);
      idx = lower.indexOf(termLower, idx + termLower.length);
    }

    if (matches.length === 0) {
      resultsEl.innerHTML = `<div class="result-count">No matches for "${escapeHtml(term)}".</div>`;
      setStatus("");
      return;
    }

    const countDiv = document.createElement("div");
    countDiv.className = "result-count";
    countDiv.textContent = `${matches.length} match${matches.length > 1 ? "es" : ""} for "${term}"`;
    resultsEl.appendChild(countDiv);

    matches.slice(0, 25).forEach((pos) => {
      const start = Math.max(0, pos - 40);
      const end = Math.min(body.length, pos + termLower.length + 40);
      const before = escapeHtml(body.slice(start, pos));
      const match = escapeHtml(body.slice(pos, pos + termLower.length));
      const after = escapeHtml(body.slice(pos + termLower.length, end));

      const snippetDiv = document.createElement("div");
      snippetDiv.className = "result-snippet";
      snippetDiv.innerHTML = `${start > 0 ? "…" : ""}${before}<mark>${match}</mark>${after}${end < body.length ? "…" : ""}`;
      resultsEl.appendChild(snippetDiv);
    });

    if (matches.length > 25) {
      const moreDiv = document.createElement("div");
      moreDiv.className = "result-count";
      moreDiv.textContent = `+ ${matches.length - 25} more matches not shown`;
      resultsEl.appendChild(moreDiv);
    }

    setStatus("");
  } catch (err) {
    setStatus("Search failed: " + err.message, "error");
  }
}

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/* ---------- Categorize tab ---------- */

function wireCategorizeTab() {
  document.getElementById("suggestCategoryBtn").addEventListener("click", suggestCategory);
}

function getMasterCategories() {
  return new Promise((resolve, reject) => {
    if (!Office.context.mailbox.masterCategories) {
      resolve([]);
      return;
    }
    Office.context.mailbox.masterCategories.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value || []);
      else reject(result.error);
    });
  });
}

function addMasterCategory(name) {
  return new Promise((resolve, reject) => {
    const colors = Object.values(Office.MailboxEnums.CategoryColor);
    const color = colors[Math.floor(Math.random() * colors.length)];
    Office.context.mailbox.masterCategories.addAsync(
      [{ displayName: name, color }],
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(result.error);
      }
    );
  });
}

function applyCategoryToItem(name) {
  return new Promise((resolve, reject) => {
    if (!currentItem.categories) {
      reject(new Error("Categories aren't available for this item."));
      return;
    }
    currentItem.categories.addAsync([name], (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(result.error);
    });
  });
}

async function suggestCategory() {
  const resultEl = document.getElementById("categoryResult");
  resultEl.innerHTML = "";

  if (isComposeMode) {
    setStatus("Categories can only be applied while reading a received or saved email.", "error");
    return;
  }

  setStatus("Analyzing this email...", "");
  try {
    const subject = currentItem.subject || "(no subject)";
    const body = await getEmailBodyText();
    const existing = await getMasterCategories();
    const existingNames = existing.map((c) => c.displayName);

    const suggestion = await callOpenAI(
      "You sort emails into short category labels (1-3 words, e.g. 'Finance', 'Travel', 'Newsletters', 'Urgent', 'Client - Acme Co'). " +
        "If one of the existing categories fits well, reuse its exact name. Otherwise propose a new short one. " +
        "Output ONLY the category name, nothing else.",
      `Existing categories: ${existingNames.join(", ") || "(none yet)"}\n\n` +
        `Email subject: ${subject}\n\nEmail body:\n${body.slice(0, 3000)}`
    );

    const clean = suggestion.replace(/["'.]/g, "").trim();
    resultEl.innerHTML = `<div>Suggested category: <span class="category-pill">${escapeHtml(clean)}</span></div>`;

    const applyBtn = document.createElement("button");
    applyBtn.textContent = `Apply "${clean}" to this email`;
    applyBtn.className = "spaced";
    applyBtn.addEventListener("click", async () => {
      applyBtn.disabled = true;
      try {
        if (!existingNames.includes(clean)) {
          await addMasterCategory(clean);
        }
        await applyCategoryToItem(clean);
        setStatus(`Applied category "${clean}".`, "success");
      } catch (err) {
        setStatus("Could not apply category: " + err.message, "error");
        applyBtn.disabled = false;
      }
    });
    resultEl.appendChild(applyBtn);
    setStatus("");
  } catch (err) {
    setStatus("Error: " + err.message, "error");
  }
}
