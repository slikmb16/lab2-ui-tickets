const groupSelect = document.getElementById("group-select");
const studentSelect = document.getElementById("student-select");
const generateBtn = document.getElementById("generate-btn");
const hint = document.getElementById("hint");
const statusLine = document.getElementById("status-line");
const modal = document.getElementById("result-modal");
const ticketTitle = document.getElementById("ticket-title");
const ticketStub = document.getElementById("ticket-stub-num");
const questionsEl = document.getElementById("questions");
const repeatBadge = document.getElementById("repeat-badge");
const toast = document.getElementById("toast");
const testList = document.getElementById("test-list");
const testsSummary = document.getElementById("tests-summary");
const progressTrack = document.getElementById("tests-progress-track");
const progressBar = document.getElementById("tests-progress-bar");

let testsCatalog = [];

function showToast(message) {
  toast.hidden = false;
  toast.textContent = message;
  clearTimeout(showToast._t);
  showToast._t = setTimeout(() => {
    toast.hidden = true;
  }, 4200);
}

function resetStudentAndButton() {
  studentSelect.value = "";
  generateBtn.disabled = true;
}

function updateGenerateEnabled() {
  generateBtn.disabled = !(groupSelect.value && studentSelect.value);
}

async function fetchJson(url, options) {
  let res;
  try {
    res = await fetch(url, options);
  } catch (_err) {
    throw new Error("Нет связи с приложением. Убедитесь, что сервер запущен, затем обновите страницу.");
  }
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data.error || `Ошибка ${res.status}`);
  }
  return data;
}

async function loadStatus() {
  const data = await fetchJson("/api/status");
  document.getElementById("stat-groups").textContent = data.groups.length;
  document.getElementById("stat-tickets").textContent = data.tickets;
  document.getElementById("stat-journal").textContent = data.journal;
  statusLine.textContent = data.fatal || data.status || "";
  statusLine.style.color = data.fatal ? "var(--danger)" : "var(--muted)";

  groupSelect.innerHTML = '<option value="">Выберите группу</option>';
  if (data.fatal) {
    groupSelect.disabled = true;
    studentSelect.disabled = true;
    generateBtn.disabled = true;
    hint.textContent = data.fatal;
    return;
  }

  data.groups.forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    groupSelect.appendChild(opt);
  });
  groupSelect.disabled = false;
}

async function onGroupChange() {
  const group = groupSelect.value;
  resetStudentAndButton();
  hint.textContent = "";
  studentSelect.innerHTML = '<option value="">Выберите студента</option>';

  if (!group) {
    studentSelect.disabled = true;
    studentSelect.innerHTML = '<option value="">Сначала выберите группу</option>';
    return;
  }

  const data = await fetchJson(`/api/groups/${encodeURIComponent(group)}/students`);
  if (data.empty) {
    studentSelect.disabled = true;
    studentSelect.innerHTML = '<option value="">В группе нет студентов</option>';
    hint.textContent = `В группе «${group}» нет студентов.`;
    return;
  }

  data.students.forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    studentSelect.appendChild(opt);
  });
  studentSelect.disabled = false;
}

function openResult(payload) {
  ticketTitle.textContent = `Ваш билет № ${payload.ticket_number}`;
  ticketStub.textContent = String(payload.ticket_number).padStart(2, "0");
  questionsEl.innerHTML = "";
  payload.questions.slice(0, 3).forEach((q) => {
    const li = document.createElement("li");
    li.className = "question";
    li.dataset.testid = "question";
    li.textContent = q;
    questionsEl.appendChild(li);
  });
  repeatBadge.hidden = !payload.repeat;
  modal.hidden = false;
}

function closeResult() {
  if (modal.hidden) return;
  modal.hidden = true;
  resetStudentAndButton();
}

async function generateTicket() {
  hint.textContent = "";
  generateBtn.disabled = true;
  generateBtn.dataset.defaultText ??= generateBtn.textContent;
  generateBtn.textContent = "Сохраняем билет…";
  try {
    const payload = await fetchJson("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        group: groupSelect.value,
        student: studentSelect.value,
      }),
    });
    openResult(payload);
    const status = await fetchJson("/api/status");
    document.getElementById("stat-journal").textContent = status.journal;
  } catch (err) {
    hint.textContent = err.message;
    showToast(err.message);
  } finally {
    generateBtn.textContent = generateBtn.dataset.defaultText;
    updateGenerateEnabled();
  }
}

async function loadJournal() {
  let data;
  try {
    data = await fetchJson("/api/results");
  } catch (err) {
    showToast(err.message);
    hint.textContent = err.message;
    return;
  }
  const body = document.getElementById("journal-body");
  if (!data.rows.length) {
    body.innerHTML = '<tr><td colspan="6" class="muted">Пока пусто — сгенерируйте первый билет.</td></tr>';
    return;
  }
  // Текст из Excel выводим через textContent, а не HTML: содержимое журнала
  // остаётся обычным текстом даже при случайных символах в исходных файлах.
  body.innerHTML = "";
  data.rows
    .forEach((r) => {
      const row = document.createElement("tr");
      [r.group, r.surname, r.name, r.ticket ? `№ ${r.ticket}` : "", r.datetime, r.repeat]
        .forEach((value) => {
          const cell = document.createElement("td");
          cell.textContent = value ?? "";
          row.appendChild(cell);
        });
      body.appendChild(row);
    });
}

function setTestRowState(id, status, message) {
  const row = document.querySelector(`[data-test-id="${CSS.escape(id)}"]`);
  if (!row) return;
  row.classList.remove("is-pass", "is-fail", "is-run");
  if (status === "passed") row.classList.add("is-pass");
  if (status === "failed") row.classList.add("is-fail");
  if (status === "running") row.classList.add("is-run");
  const label =
    status === "passed" ? "успех" : status === "failed" ? "ошибка" : status === "running" ? "идёт…" : "ожидание";
  row.querySelector(".test-status").textContent = label;
  row.querySelector(".test-result").textContent = message
    ? `Результат: ${message.split("\n")[0]}`
    : status === "passed"
      ? "Результат: проверка пройдена."
      : row.dataset.defaultHint;
}

function renderTests(catalog) {
  testsCatalog = catalog;
  testList.innerHTML = catalog
    .map(
      (t, i) => `<li class="test-row" data-test-id="${t.id}">
        <div>
          <strong>${i + 1}. ${t.label}</strong>
          <small class="test-criterion">${t.criterion}</small>
          <p class="test-description">${t.description}</p>
          <p class="test-expectation">Ожидается: ${t.expected}</p>
          <small class="test-result">Результат: ожидает запуска.</small>
        </div>
        <span class="test-status">ожидание</span>
        <button type="button" class="run-one" data-run-id="${t.id}">Запустить</button>
      </li>`
    )
    .join("");
  testList.querySelectorAll(".test-result").forEach((el) => {
    el.parentElement.parentElement.dataset.defaultHint = el.textContent;
  });
}

async function runOneTest(id) {
  setTestRowState(id, "running", "");
  const result = await fetchJson("/api/tests/run", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ id }),
  });
  const detail = result.details[0];
  if (detail) setTestRowState(detail.id, detail.status, detail.message);
  return result.ok;
}

async function runAllTests() {
  const btn = document.getElementById("run-all-tests");
  btn.disabled = true;
  progressTrack.hidden = false;
  progressBar.style.width = "0%";
  let passed = 0;
  for (let i = 0; i < testsCatalog.length; i += 1) {
    try {
      const ok = await runOneTest(testsCatalog[i].id);
      if (ok) passed += 1;
    } catch (err) {
      setTestRowState(testsCatalog[i].id, "failed", err.message);
    }
    progressBar.style.width = `${((i + 1) / testsCatalog.length) * 100}%`;
  }
  testsSummary.textContent = `Готово: ${passed} из ${testsCatalog.length} успешно.`;
  btn.disabled = false;
}

document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((t) => t.classList.remove("is-active"));
    tab.classList.add("is-active");
    document.querySelectorAll(".view").forEach((v) => v.classList.remove("is-active"));
    document.getElementById(`view-${tab.dataset.view}`).classList.add("is-active");
    if (tab.dataset.view === "journal") loadJournal();
  });
});

groupSelect.addEventListener("change", onGroupChange);
studentSelect.addEventListener("change", updateGenerateEnabled);
generateBtn.addEventListener("click", generateTicket);
document.getElementById("refresh-journal").addEventListener("click", loadJournal);
document.getElementById("refresh-sources").addEventListener("click", async () => {
  await loadStatus();
  showToast(groupSelect.disabled ? "Проверьте сообщение об ошибке: источник недоступен." : "Файлы найдены и успешно прочитаны.");
});
document.getElementById("run-all-tests").addEventListener("click", runAllTests);
testList.addEventListener("click", (e) => {
  const btn = e.target.closest("[data-run-id]");
  if (btn) runOneTest(btn.dataset.runId);
});
modal.querySelectorAll("[data-close-result]").forEach((element) => {
  element.addEventListener("click", closeResult);
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") closeResult();
});

(async function init() {
  try {
    await loadStatus();
    const tests = await fetchJson("/api/tests");
    renderTests(tests.tests);
  } catch (err) {
    hint.textContent = err.message;
  }
})();
