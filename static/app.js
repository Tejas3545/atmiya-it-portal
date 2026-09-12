/* app.js — Atmiya University FOET Attendance Portal (department-aware) */

// ─────────────────────────────────────────────────────────────────────────────
// Detect current department from the URL path
// e.g.  /it  →  "it"   /cse  →  "cse"
// ─────────────────────────────────────────────────────────────────────────────
const DEPT_META = {
  it:    { name: "B.Tech Information Technology",            short: "IT" },
  cse:   { name: "B.Tech Computer Science & Engineering",    short: "CSE" },
  civil: { name: "B.Tech Civil Engineering",                 short: "Civil" },
  mech:  { name: "B.Tech Mechanical Engineering",            short: "Mech" },
  elec:  { name: "B.Tech Electrical Engineering",            short: "Elec" },
  ce:    { name: "B.Tech Computer Engineering",              short: "CE" },
};

const pathParts = window.location.pathname.split("/").filter(Boolean);
const DEPT_KEY  = pathParts[0] ? pathParts[0].toLowerCase() : "it";
const deptMeta  = DEPT_META[DEPT_KEY] || DEPT_META["it"];

// ─────────────────────────────────────────────────────────────────────────────
// DOM References
// ─────────────────────────────────────────────────────────────────────────────
const semSelect  = document.getElementById("semester-select");
const inputEl    = document.getElementById("student-id");
const form       = document.getElementById("search-form");
const searchBtn  = document.getElementById("search-btn");
const errorBox   = document.getElementById("error-box");
const resultSec  = document.getElementById("result-section");
const searchSec  = document.getElementById("search-section");
const comingSoon = document.getElementById("coming-soon-card");

// ─────────────────────────────────────────────────────────────────────────────
// Boot — set page labels only, always show the search form
// ─────────────────────────────────────────────────────────────────────────────
(function boot() {
  const badgeEl = document.getElementById("dept-badge-label");
  if (badgeEl) badgeEl.textContent = deptMeta.name;

  const bcDept = document.getElementById("breadcrumb-dept");
  if (bcDept) bcDept.textContent = deptMeta.short;

  const pageTitle = document.getElementById("page-title");
  if (pageTitle) pageTitle.textContent = `${deptMeta.short} Attendance Portal | Atmiya University`;

  // Always show search form — never block with Coming Soon
  if (comingSoon) comingSoon.style.display = "none";
  if (searchSec)  searchSec.style.display  = "block";
})();

// ─────────────────────────────────────────────────────────────────────────────
// Form Submit
// ─────────────────────────────────────────────────────────────────────────────
form.addEventListener("submit", async (e) => {
  e.preventDefault();
  const semester = semSelect.value.trim();
  const query    = inputEl.value.trim();

  clearError();

  if (!semester) {
    showError("Please select your semester first.");
    semSelect.classList.add("error");
    return;
  }
  if (!query) {
    showError("Please enter your Registration or Enrollment number.");
    inputEl.classList.add("error");
    return;
  }

  await fetchStudent(query, semester);
});

semSelect.addEventListener("change", () => {
  semSelect.classList.remove("error");
  clearError();
});

inputEl.addEventListener("input", () => {
  inputEl.classList.remove("error");
  clearError();
});

// ─────────────────────────────────────────────────────────────────────────────
// Fetch Student
// ─────────────────────────────────────────────────────────────────────────────
async function fetchStudent(query, semester) {
  setLoading(true);
  hideResults();

  try {
    // Always call the correct dept API — URL never exposed to frontend HTML
    const url = `/api/${DEPT_KEY}/student?q=${encodeURIComponent(query)}&semester=${encodeURIComponent(semester)}`;
    const res  = await fetch(url);
    const data = await res.json();

    if (!res.ok) {
      showError(data.error || "Student not found.");
      return;
    }

    render(data);
  } catch {
    showError("Cannot connect to server. Please check your internet connection.");
  } finally {
    setLoading(false);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Render Result
// ─────────────────────────────────────────────────────────────────────────────
function render(d) {
  /* Student Identity */
  setText("info-name",      d.name || "N/A");
  setText("info-roll",      d.roll_no || "N/A");
  setText("info-enr",       d.enr_no || "—");
  setText("info-reg",       d.reg_no || "N/A");
  setText("info-div-batch", `${d.div || "N/A"} / ${d.batch || "N/A"}`);

  /* Points */
  const pts   = d.points ?? 0;
  const ptsEl = document.getElementById("stat-points");
  if (ptsEl) {
    ptsEl.textContent = pts >= 0 ? `+${pts}` : `${pts}`;
    ptsEl.style.color = pts < 0 ? "var(--red-primary)" : pts > 0 ? "var(--green-primary)" : "var(--text-main)";
  }

  /* Status Badge */
  const tag = document.getElementById("status-tag");
  const st  = (d.status || "").toLowerCase();
  if (tag) {
    tag.textContent = d.status || "N/A";
    tag.className   = "status-badge";
    if (st === "ok")        tag.classList.add("ok");
    else if (st === "pending")   tag.classList.add("pending");
    else if (st === "completed") tag.classList.add("completed");
  }

  /* Metrics */
  setText("stat-attended",  d.total_hours_attended ?? 0);
  setText("stat-of-total",  `of ${d.total_hours_required || 0} hrs required`);
  setText("stat-comp-req",  d.comp_required ?? 0);
  setText("stat-comp-done", d.comp_completed ?? 0);

  /* Donut Chart */
  const pct  = d.attendance_pct ?? 0;
  const circ = 2 * Math.PI * 40;
  const filled = (pct / 100) * circ;
  const arc  = document.getElementById("donut-arc");
  if (arc) {
    arc.style.strokeDasharray = `${filled} ${circ - filled}`;
    arc.className = "donut-arc";
    if (pct < 60) arc.classList.add("pend");
    else if (pct < 80) arc.classList.add("warn");
  }

  const donutVal = document.getElementById("donut-pct");
  if (donutVal) {
    donutVal.textContent = `${pct.toFixed(1)}%`;
    donutVal.style.color = pct < 60 ? "var(--red-primary)" : pct < 80 ? "var(--orange-primary)" : "var(--green-primary)";
  }

  /* Attendance Summary */
  const attended  = d.total_hours_attended ?? 0;
  const required  = d.total_hours_required ?? 0;
  const remaining = Math.max(0, required - attended);

  setText("sum-attended", attended);
  setText("sum-required", required);
  setText("sum-pct",      `${pct.toFixed(2)}%`);
  setText("leg-att",      attended);
  setText("leg-rem",      remaining);

  /* Compensation Progress */
  const compReq  = d.comp_required ?? 0;
  const compDone = d.comp_completed ?? 0;
  const compPct  = compReq > 0 ? Math.min(100, (compDone / compReq) * 100) : (compReq === 0 ? 100 : 0);

  setText("comp-pct-big", `${compPct.toFixed(2)}%`);
  const compFill = document.getElementById("comp-fill");
  if (compFill) compFill.style.width = `${compPct}%`;

  setText("comp-done-num", `${compDone} Hours`);
  setText("comp-req-num",  `${compReq} Hours`);

  /* Weekly Grid */
  const grid = document.getElementById("week-grid");
  if (grid) {
    grid.innerHTML = "";
    const weeks = d.weeks || [];
    weeks.forEach((w) => {
      const div = document.createElement("div");
      let cls, pctStr, badgeText;

      if (w.attendance === null) {
        cls       = "na";
        pctStr    = "—";
        badgeText = "N/A";
      } else {
        cls       = w.status.toLowerCase();
        pctStr    = `${w.attendance.toFixed(2)}%`;
        badgeText = w.status === "NOT_AVAILABLE" ? "N/A" : w.status;
      }

      div.className = `week-box ${cls}`;
      div.innerHTML = `
        <span class="week-title">Week ${w.week}</span>
        <span class="week-pct-num">${pctStr}</span>
        <span class="week-status-pill ${cls}">${badgeText}</span>
      `;
      grid.appendChild(div);
    });
  }

  showResults();
}

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────
function setText(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = val ?? "N/A";
}

function setLoading(on) {
  searchBtn.classList.toggle("loading", on);
  searchBtn.disabled = on;
  inputEl.disabled   = on;
  semSelect.disabled = on;
}

function showError(msg) {
  errorBox.textContent = msg;
  errorBox.classList.add("visible");
}

function clearError() {
  errorBox.textContent = "";
  errorBox.classList.remove("visible");
}

function showResults() {
  resultSec.classList.add("visible");
  const yOffset = -80;
  const y = resultSec.getBoundingClientRect().top + window.pageYOffset + yOffset;
  window.scrollTo({ top: Math.max(0, y), behavior: "smooth" });
}

function hideResults() {
  resultSec.classList.remove("visible");
}

function resetSearch() {
  hideResults();
  clearError();
  inputEl.value   = "";
  semSelect.value = "";
  inputEl.classList.remove("error");
  semSelect.classList.remove("error");
  window.scrollTo({ top: 0, behavior: "smooth" });
  semSelect.focus();
}

// ─────────────────────────────────────────────────────────────────────────────
// SYNC / LIVE REFRESH
// ─────────────────────────────────────────────────────────────────────────────
let _lastQuery = null;
let _lastSem   = null;

form.addEventListener("submit", () => {
  _lastQuery = inputEl.value.trim();
  _lastSem   = semSelect.value.trim();
});

const syncBtn = document.getElementById("sync-btn");

function setSyncState(on) {
  if (!syncBtn) return;
  syncBtn.classList.toggle("syncing", on);
  syncBtn.disabled = on;
  const txt = syncBtn.querySelector(".sync-btn-text");
  if (txt) txt.textContent = on ? "Refreshing..." : "Refresh";
}

async function doSync(silent = false) {
  if (!silent) setSyncState(true);
  try {
    await fetch(`/api/${DEPT_KEY}/sync`, { method: "POST" });
    if (_lastQuery && _lastSem && resultSec.classList.contains("visible")) {
      await fetchStudent(_lastQuery, _lastSem);
    }
  } catch (_) { /* silent */ }
  finally {
    if (!silent) setSyncState(false);
  }
}

async function triggerSync() { await doSync(false); }

// Auto background sync every 60 seconds (silent)
setInterval(() => doSync(true), 60_000);
