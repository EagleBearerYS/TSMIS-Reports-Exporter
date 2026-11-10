// Compare tab ▸ "Three Days Matrix" (v0.52.0) — split like the other ui-*.js, same
// global scope. A MATRIX: rows = the report editions the three-day comparison
// serves (the Clean Road files), columns = the SETS of three exported days the
// user adds, each cell = that report's three exports classified cell by cell.
// The headline is the cells that changed on the NEWEST day; under it, the cells
// that changed on the middle day and stayed, the ones that changed and went
// back, the ones that changed twice, and the records added and removed. The
// workbook lists every changed cell with its value on each day. Engine:
// three_day_matrix.

let TDM = null;                // the last three_day_matrix_info payload
let _tdRenderSeq = 0;
const _tdPick = { a: "", b: "", c: "" };   // the three pickers' current days
let _tdPickSource = null;                  // the source those picks were made for

// The matrix's OWN persisted live-formulas option (like every matrix's).
function syncThreeDayMatrixFormulas() {
  syncFormulasToggle("threeDayMatrixFormulas", "three_day_matrix_formulas");
}

function tdNum(n) {
  return Number.isFinite(n) ? n.toLocaleString() : "?";
}

// A set's days for a narrow header: month-day when the three share a year.
function tdShortDays(days) {
  const sameYear = days.every((d) => d.slice(0, 4) === days[0].slice(0, 4));
  return days.map((d) => (sameYear ? d.slice(5) : d));
}

function tdSetLabel(days) {
  return days.join(" → ");
}

// {cls, main, sub, title?, warn?} for one cell from its `cmp` state.
function tdCellContent(cmp, days) {
  cmp = cmp || {};
  const attempt = mxAttemptNote(cmp);
  if (attempt) {
    const inner = tdCellContent({ ...cmp, last_attempt: null }, days);
    return { cls: inner.cls, main: inner.main, warn: attempt.warn,
             sub: inner.sub ? `${inner.sub} · ${attempt.note}` : attempt.note,
             title: attempt.why || inner.title };
  }
  const short = tdShortDays(days);
  const missing = cmp.missing_days || [];
  if (missing.length || cmp.missing_side) {
    const which = missing.map((d) => short[days.indexOf(d)] || d).join(", ");
    return { cls: "mx-missing", main: "needs export",
             sub: which ? `not exported ${which}` : "not exported" };
  }
  if (!cmp.built) return { cls: "mx-stale", main: "compare", sub: "not built yet" };
  const t = cmp.three_day;
  if (!t) return { cls: "mx-stale", main: "re-run", sub: "outcome unknown" };
  const main = `${tdNum(t.newest)} changed ${short[2]}`;
  const parts = [`${tdNum(t.middle)} changed ${short[1]}`];
  if (t.back) parts.push(`${tdNum(t.back)} back`);
  if (t.both) parts.push(`${tdNum(t.both)} twice`);
  if (t.added) parts.push(`+${tdNum(t.added)}`);
  if (t.removed) parts.push(`−${tdNum(t.removed)}`);
  if (t.middle_only) parts.push(`${tdNum(t.middle_only)} ${short[1]} only`);
  const pairs = t.pairs || {};
  const pairLine = (key, a, b) => {
    const p = pairs[key];
    if (!p) return "";
    return `\n  ${a} vs ${b}: ${tdNum(p.differing_cells)} differing cells · `
      + `${tdNum(p.only_earlier + p.only_later)} one-sided`;
  };
  const rec = t.records || {};
  const title = `Changed on ${days[2]} (same on the two earlier days): ${tdNum(t.newest)} cells\n`
    + `Changed on ${days[1]} and still so on ${days[2]}: ${tdNum(t.middle)}\n`
    + `Changed on ${days[1]}, back on ${days[2]}: ${tdNum(t.back)}\n`
    + `Changed on both days: ${tdNum(t.both)}\n`
    + `Records on all three days: ${tdNum(rec.all)} · added: ${tdNum(t.added)} · `
    + `removed: ${tdNum(t.removed)} · on or off ${days[1]} only: ${tdNum(t.middle_only)}\n`
    + "As two-day comparisons:"
    + pairLine("1-2", days[0], days[1]) + pairLine("2-3", days[1], days[2])
    + pairLine("1-3", days[0], days[2]);
  if (cmp.completion === "partial")
    return { cls: "mx-partial", main: "partial", sub: `${main} observed — refresh`, title };
  if (cmp.stale) return { cls: "mx-stale", main, sub: "stale — refresh", title };
  const moved = (t.changed || 0) + (t.two_day || 0) + (t.added || 0) + (t.removed || 0)
    + (t.middle_only || 0);
  if (moved === 0)
    return { cls: "mx-match", main: "✓ no change", sub: "the same on all three days", title };
  return { cls: moved >= MX_HI ? "mx-diff-hi" : "mx-diff-lo", main, sub: parts.join(" · "), title };
}

// The three day pickers: any three exported days, in any order.
function renderThreeDayPickers(snap, locked) {
  const avail = snap.available_days || [];           // newest first
  const have = new Set((snap.sets || []).map((s) => s.id));
  const sels = { a: $("threeDayPickA"), b: $("threeDayPickB"), c: $("threeDayPickC") };
  // First paint, and every change of source: the three newest days, oldest on
  // top. (A pick carried over from another source could leave two pickers on
  // one day.) Otherwise a pick stays while its day is still offered.
  const defaults = { a: avail[2] || "", b: avail[1] || "", c: avail[0] || "" };
  if (_tdPickSource !== snap.source) {
    _tdPickSource = snap.source;
    Object.assign(_tdPick, defaults);
  }
  Object.entries(sels).forEach(([k, sel]) => {
    if (!avail.includes(_tdPick[k])) _tdPick[k] = defaults[k];
    sel.textContent = "";
    if (!avail.length) {
      const o = document.createElement("option");
      o.value = ""; o.textContent = "— no exported days —";
      sel.appendChild(o);
    }
    avail.forEach((d) => {
      const o = document.createElement("option");
      o.value = d; o.textContent = mxDayOptionText(d, snap);
      if (d === _tdPick[k]) o.selected = true;
      sel.appendChild(o);
    });
    sel.disabled = locked || avail.length < 3;
    sel.onchange = () => { _tdPick[k] = sel.value; syncThreeDayPick(); };
  });
  const btn = $("btnThreeDayAddSet");
  btn.onclick = () => mxBusyClick(btn, async () => {
    const r = await api.add_three_day_matrix_set(_tdPick.a, _tdPick.b, _tdPick.c);
    if (r && r.error) showMessage("error", "Can't add these days", r.error);
    await renderThreeDayMatrix();
  });
  syncThreeDayPick(have, avail.length);
}

// The line under the pickers says what the three picks will compare, and the
// Add button is live only for three different days not already a column.
function syncThreeDayPick(have, nAvail) {
  const snap = TDM || {};
  have = have || new Set((snap.sets || []).map((s) => s.id));
  nAvail = nAvail == null ? (snap.available_days || []).length : nAvail;
  const hint = $("threeDayPickHint"), btn = $("btnThreeDayAddSet");
  const locked = !!(S.st && S.st.task);
  const picks = [_tdPick.a, _tdPick.b, _tdPick.c].filter(Boolean);
  const days = [...new Set(picks)].sort();
  let ok = false, text;
  if (nAvail < 3) {
    text = "This source needs at least three exported days.";
  } else if (days.length < 3) {
    text = "Pick three different days.";
  } else if (have.has(days.join("_"))) {
    text = `${tdSetLabel(days)} is already a column.`;
  } else {
    ok = true;
    text = `Will compare ${tdSetLabel(days)}.`;
  }
  if (hint) hint.textContent = text;
  if (btn) { btn.disabled = locked || !ok; btn.dataset.ready = ok ? "1" : ""; }
}

async function renderThreeDayMatrix() {
  const grid = $("threeDayMatrixGrid");
  if (!grid) return;
  const seq = ++_tdRenderSeq;
  let snap;
  try { snap = await api.three_day_matrix_info(); } catch (e) { return; }
  if (seq !== _tdRenderSeq || !snap) return;     // a newer render started
  if (snap.error) return;
  TDM = snap;
  const sets = snap.sets || [], locked = !!(S.st && S.st.task);

  const srcSel = $("threeDayMatrixSource");
  srcSel.textContent = "";
  (snap.sources || []).forEach((s) => {
    const o = document.createElement("option");
    o.value = s.key; o.textContent = s.label;
    if (s.key === snap.source) o.selected = true;
    srcSel.appendChild(o);
  });
  srcSel.disabled = locked;
  srcSel.onchange = async () => {
    const r = await api.set_three_day_matrix_source(srcSel.value);
    if (r && r.error) showMessage("error", "Can't set source", r.error);
    await renderThreeDayMatrix();
  };

  renderThreeDayPickers(snap, locked);

  const rtog = $("threeDayMatrixReportToggles");
  rtog.textContent = "";
  const hidden = new Set(snap.hidden || []);
  (snap.all_rows || []).forEach((r) => {
    const isOn = !hidden.has(r.key);
    rtog.appendChild(mxToggleChip(r.label, isOn, (isOn ? "Hide " : "Show ") + r.label, locked,
      (next) => api.set_three_day_matrix_report(r.key, next), renderThreeDayMatrix));
  });

  grid.textContent = "";
  if (!sets.length) {
    grid.style.gridTemplateColumns = ""; grid.style.gridTemplateRows = "";
    const empty = document.createElement("div");
    empty.className = "dm-empty";
    empty.textContent = "Pick three exported days in Matrix options and add them. Each "
      + "report is then compared across the three: what changed on the newest day, what "
      + "changed on the middle day and stayed, and what changed and went back.";
    grid.appendChild(empty);
    wireThreeDayFooter();
    updateThreeDayMatrixProgress();
    return;
  }
  grid.style.gridTemplateColumns = `minmax(210px,1.1fr) repeat(${sets.length}, minmax(190px,1fr))`;
  grid.style.gridTemplateRows = ["auto"].concat(snap.rows.map(() => "minmax(54px,1fr)")).join(" ");

  const corner = document.createElement("div");
  corner.className = "mx-cell mx-corner mx-colhead";
  corner.textContent = "Report \\ Three days";
  grid.appendChild(corner);
  sets.forEach((s) => {
    const days = s.days, full = tdSetLabel(days);
    const h = document.createElement("div");
    h.className = "mx-cell mx-colhead";
    h.title = full;
    const lab = document.createElement("div");
    lab.textContent = tdShortDays(days).join(" → ");
    h.appendChild(lab);
    const year = document.createElement("div");
    year.className = "mx-sub mx-daysite";
    const site = mxDaySite(days[2], snap);
    year.textContent = [...new Set(days.map((d) => d.slice(0, 4)))].join(" – ")
      + (site ? ` · ${site}` : "");
    h.appendChild(year);
    const btns = document.createElement("span"); btns.className = "mxch-btns";
    btns.append(
      mxHeadBtn("i-compare", `Rebuild every report for ${full}`, "mxch-rebuild",
        async () => {
          const r = await api.rebuild_three_day_matrix("all", null, s.id);
          if (r && r.nothing) showMessage("info", "Nothing to rebuild", "No report has all three of these days exported.");
          else if (r && r.error) showMessage("error", "Can't rebuild", r.error);
        }),
      mxHeadBtn("i-trash", `Remove the ${full} column`, "mxch-rm", async () => {
        await api.remove_three_day_matrix_set(s.id); await renderThreeDayMatrix();
      }));
    h.appendChild(btns);
    grid.appendChild(h);
  });

  snap.rows.forEach((rk) => {
    const rlabel = snap.row_labels[rk] || rk;
    const rh = document.createElement("div");
    rh.className = "mx-cell mx-rowhead";
    rh.dataset.rk = rk; rh.dataset.label = rlabel;
    const top = document.createElement("div"); top.className = "mxrh-top";
    const lbl = document.createElement("span"); lbl.className = "mxrh-label";
    lbl.textContent = rlabel;
    top.appendChild(lbl);
    top.appendChild(mxHeadBtn("i-compare", `Rebuild ${rlabel} for every set of days`,
      "mxch-rebuild", async () => {
        const r = await api.rebuild_three_day_matrix("all", rk, null);
        if (r && r.nothing) showMessage("info", "Nothing to rebuild", "No set of days has all three exports of this report.");
        else if (r && r.error) showMessage("error", "Can't rebuild", r.error);
      }));
    rh.appendChild(top);
    dndAttach(rh, top, rk, "td-row", "y", () => snap.rows.slice(), async (order) => {
      const r = await api.set_three_day_matrix_row_order(order);
      if (r && r.error) showMessage("error", "Can't reorder", r.error);
      else await renderThreeDayMatrix();
    });
    grid.appendChild(rh);
    sets.forEach((s) => {
      const c = (snap.cells[rk] || {})[s.id] || {}, cmp = c.cmp || {};
      const cell = document.createElement("div"); cell.className = "mx-cell";
      const main = document.createElement("div"); main.className = "mx-num";
      const sub = document.createElement("div"); sub.className = "mx-sub";
      const v = tdCellContent(cmp, s.days);
      cell.classList.add(v.cls); main.textContent = v.main; sub.textContent = v.sub;
      if (v.warn) cell.classList.add(v.warn);
      const exports = (c.exports || []).map((e) =>
        `\n  ${e.date}: ${e.present ? "exported " + fmtAge(e.age_seconds) : "not exported"}`).join("");
      cell.title = `${rlabel} — ${tdSetLabel(s.days)}${exports}`
        + (v.title ? `\n\n${v.title}` : "");
      cell.append(main, sub);
      const acts = document.createElement("div"); acts.className = "mx-actions";
      if (!cmp.missing_side) {
        acts.appendChild(mxActBtn("i-compare", "Build / rebuild this three-day comparison", false,
          async () => {
            const r = await api.build_three_day_matrix_cell(rk, s.id);
            if (r && r.error) showMessage("error", "Can't build", r.error);
          }));
      }
      if (cmp.built) {
        const ob = mxActBtn("i-external", "Open this three-day workbook", false, async () => {
          const r = await api.open_three_day_cell_comparison(rk, s.id);
          if (r && r.error) showMessage("error", "Can't open", r.error);
        });
        ob.classList.add("mx-open"); acts.appendChild(ob);
      }
      cell.appendChild(acts);
      grid.appendChild(cell);
    });
  });

  wireThreeDayFooter();
  syncThreeDayMatrixFormulas();
  updateThreeDayMatrixProgress();
}

function wireThreeDayFooter() {
  $("btnThreeDayBuildAll").onclick = async () => {
    const r = await api.rebuild_three_day_matrix("all");
    if (r && r.nothing) showMessage("info", "Nothing to compare", "Add three exported days first: a cell needs the report exported on all three.");
    else if (r && r.error) showMessage("error", "Can't compare", r.error);
  };
  $("btnThreeDayRebuildAll").onclick = async () => {
    const r = await api.rebuild_three_day_matrix("stale");
    if (r && r.nothing) showMessage("info", "Up to date", "Every three-day comparison is current.");
    else if (r && r.error) showMessage("error", "Can't rebuild", r.error);
  };
  $("btnOpenThreeDayComparisons").onclick = async () => {
    const r = await api.open_three_day_comparisons_folder();
    if (r && r.error) showMessage("error", "Can't open", r.error);
  };
  $("btnThreeDayCancel").onclick = () => api.cancel_run();
}

// Called on every state push while the sub-tab shows (no API call): the live
// progress line, the Cancel button, the queue panel and the lock-sensitive bits.
function updateThreeDayMatrixProgress() {
  const el = $("threeDayMatrixProgress");
  if (el) {
    const m = S.st && S.st.matrix;
    if (m && m.total) { el.hidden = false; el.textContent = mxProgressText(m); }
    else el.hidden = true;
  }
  const locked = !!(S.st && S.st.task);
  document.querySelectorAll(
    "#threeDayMatrixSource, #threeDayMatrixReportToggles .mx-toggle, "
    + "#threeDayMatrixGrid .mxrh-top .mxch-rebuild")
    .forEach((c) => { c.disabled = locked; });
  const few = !TDM || (TDM.available_days || []).length < 3;
  ["threeDayPickA", "threeDayPickB", "threeDayPickC"].forEach((id) => {
    const sel = $(id);
    if (sel) sel.disabled = locked || few;
  });
  const add = $("btnThreeDayAddSet");
  if (add) add.disabled = locked || add.dataset.ready !== "1";
  const cancel = $("btnThreeDayCancel");
  if (cancel) {
    const running = !!(S.st && S.st.task === "matrix");
    cancel.classList.toggle("hidden", !running);
    cancel.disabled = !running;
  }
  renderQueuePanel("threeDayQueueGroup", "threeDayQueue", "threeDayQueueCount");
}
