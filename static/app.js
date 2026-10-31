"use strict";

const $ = (id) => document.getElementById(id);
const state = {
  files: [], questions: [], answers: {}, graded: {}, revealed: {},
  filter: "all", warnings: [], showAll: false,
};

function node(tag, className, text) {
  const n = document.createElement(tag);
  if (className) n.className = className;
  if (text !== undefined) n.textContent = text;
  return n;
}
function toast(message) {
  const t = $("toast");
  t.textContent = message;
  t.classList.add("show");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => t.classList.remove("show"), 2800);
}
function formatBytes(bytes) {
  return bytes < 1024 * 1024 ? `${Math.ceil(bytes / 1024)}KB` : `${(bytes / 1024 / 1024).toFixed(1)}MB`;
}
function addFiles(items) {
  const accepted = [".pdf", ".pptx", ".docx", ".txt"];
  for (const f of Array.from(items)) {
    const ext = "." + (f.name.split(".").pop() || "").toLowerCase();
    if (!accepted.includes(ext)) { toast(`${f.name}: 지원하지 않는 형식이에요.`); continue; }
    if (f.size > 12 * 1024 * 1024) { toast(`${f.name}: 12MB 이하만 가능해요.`); continue; }
    if (state.files.length >= 6) { toast("파일은 최대 6개까지 가능해요."); break; }
    if (state.files.some(x => x.name === f.name && x.size === f.size)) continue;
    state.files.push(f);
  }
  renderFiles();
}
function renderFiles() {
  const list = $("fileList");
  list.replaceChildren();
  state.files.forEach((f, i) => {
    const row = node("div", "file-pill");
    row.append(node("span", "", "▤"), node("span", "file-name", f.name), node("span", "file-size", formatBytes(f.size)));
    const del = node("button", "", "×");
    del.type = "button";
    del.setAttribute("aria-label", `${f.name} 삭제`);
    del.addEventListener("click", () => { state.files.splice(i, 1); renderFiles(); });
    row.append(del);
    list.append(row);
  });
}

const dropZone = $("dropZone");
dropZone.addEventListener("click", () => $("fileInput").click());
dropZone.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("fileInput").click(); } });
$("fileInput").addEventListener("change", e => { addFiles(e.target.files); e.target.value = ""; });
for (const event of ["dragenter", "dragover"]) dropZone.addEventListener(event, e => { e.preventDefault(); dropZone.classList.add("dragging"); });
for (const event of ["dragleave", "drop"]) dropZone.addEventListener(event, e => { e.preventDefault(); dropZone.classList.remove("dragging"); });
dropZone.addEventListener("drop", e => addFiles(e.dataTransfer.files));

function setStatus(message, kind = "") {
  $("formStatus").textContent = message;
  $("formStatus").className = "form-status " + kind;
}

$("generateForm").addEventListener("submit", async e => {
  e.preventDefault();
  if (!state.files.length && !$("pastText").value.trim()) {
    setStatus("강의자료 또는 정답이 포함된 기출문제를 입력해 주세요.", "error");
    return;
  }
  const form = new FormData();
  state.files.forEach(f => form.append("files", f, f.name));
  form.append("past_text", $("pastText").value);
  form.append("concepts_text", $("conceptsText").value);
  form.append("count", $("questionCount").value);
  form.append("question_style", $("questionStyle").value);
  form.append("variants", $("variants").checked ? "true" : "false");
  const btn = $("generateBtn");
  btn.disabled = true;
  btn.textContent = "강의노트를 읽고 문제를 만들고 있어요…";
  setStatus("강의자료에서 문장을 추출하고, 기출과 정답을 대조하는 중이에요. 최대 1~2분 걸릴 수 있어요.", "loading");
  try {
    const response = await fetch("/api/generate", { method: "POST", body: form });
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || "문제를 생성하지 못했어요.");
    loadBank(result.questions || [], result.warnings || [], "AI 생성 문제은행");
    setStatus(`${result.questions.length}문제가 준비됐어요. 아래에서 바로 풀어보세요.`);
  } catch (err) {
    setStatus(err.message || "오류가 발생했어요.", "error");
  } finally {
    btn.disabled = false;
    btn.replaceChildren(node("span", "btn-star", "✳"), document.createTextNode(" AI 문제은행 생성하기 →"));
  }
});

const demo = [
  {question:"TBL은 뭐의 약자인가?",answer:"Triple Bottom Line",alternatives:[],kind:"short",options:[],priority:"past",source_label:"데모 · 기출 입력",evidence:"",verified:false,verification_note:"데모용 예시입니다."},
  {question:"2030년까지 이행하는 국제사회 공동목표의 주 목표와 세부목표 개수를 쓰시오.",answer:"17개, 169개",alternatives:["17,169"],kind:"short",options:[],priority:"extra",source_label:"데모 · 예시 강의노트",evidence:"17개 목표와 169개 세부목표",verified:false,verification_note:"데모용 예시입니다."},
  {question:"제품의 설계 단계에서 환경영향을 최소화하는 방법은 (   )이다.",answer:"친환경설계",alternatives:["DfE"],kind:"blank",options:[],priority:"keyword",source_label:"데모 · 예시 강의노트",evidence:"제품 설계 단계에서 환경영향을 최소화",verified:false,verification_note:"데모용 예시입니다."},
  {question:"기업의 사회적 책임을 뜻하는 약자는?",answer:"CSR",alternatives:[],kind:"short",options:[],priority:"extra",source_label:"데모 · 예시 강의노트",evidence:"기업의 사회적 책임",verified:false,verification_note:"데모용 예시입니다."},
  {question:"기업의 사회적 책임을 경제적·법적·윤리적·자선적 책임으로 구분한 모형은?",answer:"캐롤의 4가지 책임모형",alternatives:["캐롤 모형"],kind:"choice",options:["포터 가설","캐롤의 4가지 책임모형","지속가능성과곡선","이해관계자 맵"],priority:"extra",source_label:"데모 · 예시 강의노트",evidence:"경제적·법적·윤리적·자선적 책임",verified:false,verification_note:"데모용 예시입니다."},
];
$("demoBtn").addEventListener("click", () => loadBank(demo, ["샘플 문제은행입니다. 실제 강의자료 분석이나 AI 생성은 수행하지 않았어요."], "샘플 문제은행"));

function loadBank(questions, warnings = [], subtitle = "나의 문제은행") {
  if (!Array.isArray(questions) || !questions.length) { toast("생성된 문제가 없어요. 강의자료와 필수 개념을 확인해 주세요."); return; }
  state.questions = questions.map((q, i) => ({ ...q, _id: i }));
  state.answers = {}; state.graded = {}; state.revealed = {}; state.filter = "all"; state.showAll = false;
  state.warnings = warnings;
  $("quizSubtitle").textContent = subtitle + " · 정답 확인 후 출처를 비교해 보세요.";
  $("quizSection").hidden = false;
  $("showAllBtn").textContent = "정답 모두 보기";
  setFilter("all");
  $("quizSection").scrollIntoView({ behavior: "smooth", block: "start" });
  toast(`${questions.length}문제 준비 완료!`);
}

function clean(s) {
  return String(s ?? "").normalize("NFKC").toLowerCase().replace(/[\s·,，.()\-\/_→]/g, "");
}
function isCorrect(id) {
  const q = state.questions.find(x => x._id === id);
  const ans = clean(state.answers[id]);
  if (!ans || !q) return false;
  return [q.answer, ...(q.alternatives || [])].some(v => clean(v) === ans);
}
function markQuestion(id) {
  const v = (state.answers[id] || "").trim();
  if (!v) { toast("먼저 정답을 입력해 주세요."); return; }
  state.graded[id] = isCorrect(id) ? "correct" : "wrong";
  rerender();
}
function rerender() {
  const y = window.scrollY;
  renderQuiz();
  window.scrollTo({top:y,behavior:"instant"});
}
function stats() {
  const graded = Object.keys(state.graded).length;
  const right = Object.values(state.graded).filter(v => v === "correct" || v === "manual").length;
  $("totalStat").textContent = state.questions.length;
  $("gradedStat").textContent = graded;
  $("correctStat").textContent = right;
  $("percentStat").textContent = graded ? `${Math.round(right / graded * 100)}%` : "—";
}
function renderQuiz() {
  stats();
  const warningBox = $("warningBox");
  warningBox.hidden = !state.warnings.length;
  warningBox.textContent = state.warnings.join("\n");
  document.querySelectorAll("[data-filter]").forEach(btn => btn.classList.toggle("active", btn.dataset.filter === state.filter));
  const list = $("questionList");
  list.replaceChildren();
  const visible = state.questions.filter(q => state.filter === "all" || (state.filter === "wrong" && state.graded[q._id] === "wrong") || (state.filter === "pending" && !state.graded[q._id]));
  if (!visible.length) { list.append(node("div", "empty-state", state.filter === "wrong" ? "틀린 문제가 없어요. 잘하고 있어요!" : "해당하는 문제가 없어요.")); return; }
  visible.forEach(q => {
    const id = q._id;
    const card = node("article", "question-card");
    const meta = node("div", "question-meta");
    meta.append(node("span", "q-number", `QUESTION ${String(state.questions.indexOf(q) + 1).padStart(2, "0")}`));
    const tags = node("div", "meta-right");
    if (q.priority === "past") tags.append(node("span", "chip", "기출 원문"));
    if (q.priority === "keyword") tags.append(node("span", "chip", "필수 개념"));
    tags.append(node("span", q.verified ? "chip verified" : "chip review", q.verified ? "인용문 확인" : "검토 필요"));
    meta.append(tags); card.append(meta);
    card.append(node("p", "q-text", q.question));
    const line = node("div", "answer-line");
    if (q.kind === "choice" && Array.isArray(q.options) && q.options.length === 4) {
      const choices = node("div", "choices");
      q.options.forEach((option, i) => {
        const label = node("label", "choice");
        const radio = node("input"); radio.type = "radio"; radio.name = `choice-${id}`; radio.value = option;
        radio.checked = state.answers[id] === option;
        radio.addEventListener("change", () => { state.answers[id] = option; delete state.graded[id]; stats(); });
        label.append(radio, node("span", "", `${i+1}. ${option}`)); choices.append(label);
      });
      card.append(choices);
    } else {
      const input = node("input"); input.type = "text"; input.placeholder = "정답을 입력해 보세요";
      input.value = state.answers[id] || ""; input.setAttribute("aria-label", `문제 ${id+1} 정답`);
      input.addEventListener("input", e => { state.answers[id] = e.target.value; delete state.graded[id]; stats(); });
      input.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); markQuestion(id); } });
      line.append(input);
    }
    const check = node("button", "btn btn-ghost btn-small", "채점"); check.type = "button";
    check.addEventListener("click", () => markQuestion(id));
    line.append(check); card.append(line);
    if (state.graded[id]) {
      const ok = state.graded[id] !== "wrong";
      card.append(node("p", `question-feedback ${ok ? "correct" : "wrong"}`, ok ? "✓ 정답으로 표시했어요" : "✕ 정답을 다시 확인해 보세요"));
    }
    const footer = node("div", "question-footer");
    const toggle = node("button", "link-button", state.revealed[id] || state.showAll ? "정답 숨기기 ↑" : "정답 확인 ↓");
    toggle.type = "button";
    toggle.addEventListener("click", () => { state.revealed[id] = !(state.revealed[id] || state.showAll); state.showAll = false; rerender(); });
    footer.append(toggle);
    if (state.graded[id] === "wrong") {
      const manual = node("button", "link-button", "내 답도 정답으로 인정");
      manual.type = "button"; manual.addEventListener("click", () => { state.graded[id] = "manual"; rerender(); });
      footer.append(manual);
    }
    card.append(footer);
    if (state.revealed[id] || state.showAll) {
      const answer = node("div", "answer-reveal");
      answer.append(node("span", "", "정답  "), node("strong", "", q.answer));
      const source = node("div", "source-detail");
      source.append(node("div", "", `출처: ${q.source_label || "출처 미확인"}`));
      if (q.evidence) source.append(node("p", "", `인용: “${q.evidence}”`));
      if (q.verification_note) source.append(node("p", "", q.verification_note));
      answer.append(source); card.append(answer);
    }
    list.append(card);
  });
}
function setFilter(v) { state.filter = v; renderQuiz(); }
document.querySelectorAll("[data-filter]").forEach(btn => btn.addEventListener("click", () => setFilter(btn.dataset.filter)));
$("gradeAllBtn").addEventListener("click", () => {
  let count = 0;
  state.questions.forEach(q => { const id = q._id; if ((state.answers[id] || "").trim()) { state.graded[id] = isCorrect(id) ? "correct" : "wrong"; count++; } });
  rerender(); toast(`${count}개 답안을 채점했어요.`);
});
$("clearAnswersBtn").addEventListener("click", () => {
  if (!confirm("모든 답안과 채점 결과를 초기화할까요? 문제는 유지돼요.")) return;
  state.answers = {}; state.graded = {}; state.revealed = {}; state.showAll = false; setFilter("all");
});
$("shuffleBtn").addEventListener("click", () => {
  const a = state.questions;
  for (let i = a.length-1; i > 0; i--) { const j = Math.floor(Math.random()*(i+1)); [a[i],a[j]] = [a[j],a[i]]; }
  state.answers = {}; state.graded = {}; state.revealed = {}; state.showAll = false;
  setFilter("all"); toast("순서를 섞고 답안을 초기화했어요.");
});
$("showAllBtn").addEventListener("click", () => {
  state.showAll = !state.showAll;
  state.revealed = {};
  $("showAllBtn").textContent = state.showAll ? "정답 모두 숨기기" : "정답 모두 보기";
  rerender();
});
$("exportBtn").addEventListener("click", () => {
  const data = JSON.stringify({format:"munjegongbang-v1",exported_at:new Date().toISOString(),questions:state.questions.map(({_id,...q})=>q)},null,2);
  const url = URL.createObjectURL(new Blob([data],{type:"application/json"}));
  const link = document.createElement("a"); link.href = url; link.download = `문제공방_${new Date().toISOString().slice(0,10)}.json`; link.click();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
});
$("printBtn").addEventListener("click", () => window.print());
$("importInput").addEventListener("change", async e => {
  const file = e.target.files[0];
  if (!file) return;
  if (file.size > 2 * 1024 * 1024) { toast("JSON 파일은 2MB 이하만 가능해요."); return; }
  try {
    const data = JSON.parse(await file.text());
    if (data.format !== "munjegongbang-v1" || !Array.isArray(data.questions) || !data.questions.length || data.questions.length > 200) throw new Error("문제공방 JSON 형식이 아니에요.");
    const valid = data.questions.every(q => typeof q.question === "string" && typeof q.answer === "string" && q.question.length < 1200 && q.answer.length < 300);
    if (!valid) throw new Error("일부 문항 형식이 올바르지 않아요.");
    loadBank(data.questions, [], "불러온 문제은행");
  } catch (err) { toast(err.message || "파일을 읽지 못했어요."); }
  e.target.value = "";
});
