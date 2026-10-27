const form = document.getElementById("searchForm");
const keywordInput = document.getElementById("keyword");
const depthInput = document.getElementById("depth");
const submitBtn = document.getElementById("submitBtn");
const exportBtn = document.getElementById("exportBtn");
const confirmBtn = document.getElementById("confirmBtn");
const statusEl = document.getElementById("status");
const errorEl = document.getElementById("error");
const skippedEl = document.getElementById("skipped");
const picker = document.getElementById("picker");
const pickerHint = document.getElementById("pickerHint");
const pickerList = document.getElementById("pickerList");
const readyList = document.getElementById("readyList");
const result = document.getElementById("result");
const summary = document.getElementById("summary");

let lastQuery = null;
const MAX_BATCH = 15;

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const lines = splitLines(keywordInput.value);
  if (!lines.length) {
    showError("請輸入至少一家公司");
    return;
  }
  if (lines.length > MAX_BATCH) {
    showError(`一次最多查 ${MAX_BATCH} 家，請分批查詢`);
    return;
  }
  await resolveLines(lines);
});

exportBtn.addEventListener("click", exportExcel);

function splitLines(text) {
  return text.split(/[\n,，、;；]+/).map((line) => line.trim()).filter(Boolean);
}

async function resolveLines(lines) {
  setBusy(true, "正在核對公司名稱…");
  hide(picker);
  hide(result);
  hide(skippedEl);
  hide(errorEl);
  const ready = [];
  const choices = [];
  const missing = [];

  try {
    for (let index = 0; index < lines.length; index += 1) {
      statusEl.textContent = `正在核對第 ${index + 1} / ${lines.length} 家…`;
      const line = lines[index];
      const response = await fetch(`/api/search?q=${encodeURIComponent(line)}`);
      const data = await response.json();
      const rows = data.results || [];
      if (!response.ok || !rows.length) {
        missing.push(line);
        continue;
      }
      if (rows.length === 1) {
        ready.push(rows[0]);
      } else {
        choices.push({ line, found: data.found, results: rows });
      }
    }
  } catch {
    showError("網路連線中斷，請再試一次");
    setBusy(false);
    return;
  }

  setBusy(false);
  if (missing.length) {
    skippedEl.textContent = `這幾行找不到，已略過：${missing.join("、")}`;
    show(skippedEl);
  }
  if (!ready.length && !choices.length) {
    showError("沒有找到任何公司");
    return;
  }
  if (choices.length) {
    renderPicker(ready, choices);
    return;
  }
  await runQuery(uniqueIds(ready));
}

function renderPicker(ready, choices) {
  pickerHint.textContent = "下面有名稱對到多家公司，請各選一家，再開始查詢。";
  readyList.textContent = ready.length
    ? `已確認：${ready.map((company) => `${company.company_name}（${company.company_id}）`).join("、")}`
    : "";
  pickerList.innerHTML = choices.map((group, index) => `
    <fieldset>
      <legend>${esc(group.line)}${group.found > group.results.length ? `（共 ${group.found} 筆，先列 ${group.results.length} 筆）` : ""}</legend>
      ${group.results.map((company, item) => `
        <label class="choice">
          <input type="radio" name="group-${index}" value="${esc(company.company_id)}" ${item === 0 ? "checked" : ""}>
          <span>
            <strong>${esc(company.company_name)}</strong><br>
            ${esc(company.entity_type)}　統編 ${esc(company.company_id)}${company.status ? `　${esc(company.status)}` : ""}${company.representative ? `　代表人 ${esc(company.representative)}` : ""}
          </span>
        </label>
      `).join("")}
    </fieldset>
  `).join("");
  show(picker);
  confirmBtn.onclick = async () => {
    const picked = ready.slice();
    pickerList.querySelectorAll("fieldset").forEach((fieldset) => {
      const selected = fieldset.querySelector("input:checked");
      if (!selected) return;
      const name = selected.parentElement.querySelector("strong").textContent;
      picked.push({ company_id: selected.value, company_name: name });
    });
    hide(picker);
    await runQuery(uniqueIds(picked));
  };
}

async function runQuery(ids) {
  if (!ids.length) {
    showError("請至少選擇一家公司");
    return;
  }
  const depth = depthInput.value;
  setBusy(true, depth === "all"
    ? `正在查 ${ids.length} 家公司，並往上追溯…`
    : `正在查 ${ids.length} 家公司…`);
  hide(errorEl);
  try {
    const response = await fetch("/api/query", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ids, depth }),
    });
    const data = await response.json();
    if (!response.ok) {
      showError(data.error || "查詢失敗");
      hide(result);
      return;
    }
    lastQuery = { ids, depth };
    renderResult(data);
    clearStatus();
  } catch {
    showError("網路連線中斷，請再試一次");
  } finally {
    setBusy(false);
  }
}

function renderResult(data) {
  const names = (data.roots || []).map((root) => root.name).join("、");
  summary.innerHTML = `
    <b>${esc(names)}</b>
    <b>查詢 ${data.stats.queries} 家</b>
    <b>公司 ${data.stats.companies} 列</b>
    <b>人員欄 ${data.stats.people}</b>
    <b>往上 ${data.stats.levels} 層</b>
  `;
  const notes = document.getElementById("notes");
  if (data.notes && data.notes.length) {
    notes.textContent = data.notes.join("。");
    show(notes);
  } else {
    hide(notes);
  }
  document.getElementById("matrixCount").textContent =
    `${data.grid.length} 列公司。第一格是負責人，後面只排姓名`;
  document.getElementById("matrixHead").innerHTML = `
    <tr>
      <th>層級</th>
      <th>統一編號</th>
      <th>公司名稱</th>
      <th>登記現況</th>
      <th>來自查詢</th>
      ${data.people.map((name) => `<th class="person">${esc(name)}</th>`).join("")}
    </tr>
  `;
  document.getElementById("matrixBody").innerHTML = data.grid.map((row) => {
    const main = `
      <tr>
        <td><span class="level${row.level === 0 ? " target" : ""}">${levelLabel(row.level)}</span></td>
        <td>${row.company_id ? esc(row.company_id) : "—"}</td>
        <td>${esc(row.company_name)} ${findbizButton(row)}</td>
        <td>${esc(row.status || "—")}</td>
        <td>${esc((row.sources || []).join("、") || "—")}</td>
        ${row.cells.map((cell) => cell ? `<td>${esc(cell)}</td>` : `<td class="blank"></td>`).join("")}
      </tr>
    `;
    const english = (row.english_name || "").trim();
    if (!english || english === row.company_name) return main;
    const blanks = row.cells.map(() => "<td></td>").join("");
    return main + `
      <tr class="english-row">
        <td></td>
        <td></td>
        <td>${esc(english)}</td>
        <td></td>
        <td></td>
        ${blanks}
      </tr>
    `;
  }).join("");
  show(result);
}

async function exportExcel() {
  if (!lastQuery) return;
  exportBtn.disabled = true;
  exportBtn.textContent = "產生 Excel…";
  try {
    const response = await fetch("/api/export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(lastQuery),
    });
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      showError(data.error || "匯出失敗");
      return;
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filenameFrom(response) || "公司人員對照.xlsx";
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  } catch {
    showError("匯出時網路連線中斷，請再試一次");
  } finally {
    exportBtn.disabled = false;
    exportBtn.textContent = "匯出 Excel";
  }
}

function filenameFrom(response) {
  const header = response.headers.get("Content-Disposition") || "";
  const match = header.match(/filename\*=UTF-8''([^;]+)/i);
  if (!match) return "";
  try {
    return decodeURIComponent(match[1]);
  } catch {
    return "";
  }
}

function uniqueIds(companies) {
  const ids = [];
  const seen = new Set();
  companies.forEach((company) => {
    if (!company.company_id || seen.has(company.company_id)) return;
    seen.add(company.company_id);
    ids.push(company.company_id);
  });
  return ids;
}

function findbizButton(company) {
  if (!company.company_id) return "";
  return `
    <form method="post" action="https://findbiz.nat.gov.tw/fts/query/QueryCmpyDetail/queryCmpyDetail.do" target="_blank">
      <input type="hidden" name="banNo" value="${esc(company.company_id)}">
      <input type="hidden" name="brBanNo" value="">
      <input type="hidden" name="banKey" value="">
      <input type="hidden" name="estbId" value="">
      <input type="hidden" name="objectId" value="">
      <input type="hidden" name="CPage" value="">
      <input type="hidden" name="brCmpyPage" value="">
      <input type="hidden" name="eng" value="false">
      <button class="findbiz" type="submit">findbiz</button>
    </form>
  `;
}

function levelLabel(level) {
  if (level === 0) return "查詢公司";
  return `上層第 ${level} 層`;
}

function setBusy(busy, message) {
  submitBtn.disabled = busy;
  confirmBtn.disabled = busy;
  if (busy) {
    statusEl.textContent = message;
    show(statusEl);
    hide(errorEl);
  }
}

function clearStatus() {
  hide(statusEl);
}

function showError(message) {
  errorEl.textContent = message;
  show(errorEl);
  clearStatus();
}

function show(element) { element.hidden = false; }
function hide(element) { element.hidden = true; }

function esc(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}
