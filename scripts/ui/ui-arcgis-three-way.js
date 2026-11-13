// ArcGIS tab ▸ "Three-way" (v0.51.0) — split like the other ui-*.js, same global
// scope. A by-day MATRIX: rows = every export edition of a report that has all
// three sources (the Clean Road files), columns = exported days, each cell = that
// day's export vs the report's TSN dataset vs its ArcGIS build. The headline is
// "report ≠" — the cells where TSN and the layers agree and the report doesn't —
// so the expected TSN drift (the report follows the layers, TSN is frozen at the
// cutover) stays out of the way. The workbook lists every cell with the ArcGIS
// layer and field it comes from. Engine: three_way_matrix.

let TWM = null;                // the last three_way_matrix_info payload
let _twRenderSeq = 0;

function threeWayMatrixActive() {
  return S.tab === "arcgis" && agSub === "threeway";
}

function twNum(n) {
  return Number.isFinite(n) ? n.toLocaleString() : "?";
}

// {cls, main, sub, title?, warn?} for one cell from its `cmp` state.
function twCellContent(cmp) {
  cmp = cmp || {};
  const attempt = mxAttemptNote(cmp);
  if (attempt) {
    const inner = twCellContent({ ...cmp, last_attempt: null });
    return { cls: inner.cls, main: inner.main, warn: attempt.warn,
             sub: inner.sub ? `${inner.sub} · ${attempt.note}` : attempt.note,
             title: attempt.why || inner.title };
  }
  const missing = cmp.missing_sides || (cmp.missing_side ? [cmp.missing_side] : []);
  if (missing.length) {
    const names = { export: "not exported", tsn: "no TSN dataset", layers: "no ArcGIS build" };
    const main = missing.includes("export") ? "needs export"
      : missing.includes("tsn") ? "needs TSN" : "needs ArcGIS";
    return { cls: "mx-missing", main, sub: missing.map((m) => names[m] || m).join(" · ") };
  }
  if (!cmp.built) return { cls: "mx-unbuilt", main: "not built", sub: "build it to compare" };
  const tw = cmp.three_way;
  if (!tw) return { cls: "mx-stale", main: "re-run", sub: "outcome unknown" };
  const rd = tw.report, miss = tw.missing_from_report || 0;
  const main = `${twNum(rd)} report ≠`;
  const parts = [];
  if (miss) parts.push(`${twNum(miss)} missing`);
  parts.push(`${twNum(tw.tsn)} expected`);
  if (tw.arcgis) parts.push(`${twNum(tw.arcgis)} ArcGIS ≠`);
  const title = `Report differs (TSN = ArcGIS ≠ TSMIS): ${twNum(rd)} cells\n`
    + `ArcGIS differs (TSMIS = TSN ≠ ArcGIS): ${twNum(tw.arcgis)}\n`
    + `All three differ: ${twNum(tw.all)}\n`
    + `TSN differs — expected (TSMIS = ArcGIS ≠ TSN): ${twNum(tw.tsn)}\n`
    + `All three agree: ${twNum(tw.match)}`
    + (tw.undecided ? `\nUndecided (ArcGIS unplaced): ${twNum(tw.undecided)}` : "")
    + `\nRecords on all three: ${twNum((tw.records || {}).all)} · missing from the report: `
    + `${twNum(miss)} · only in the report: ${twNum(tw.only_in_report)}`;
  if (cmp.completion === "partial")
    return { cls: "mx-partial", main: "partial", sub: `${main} observed — refresh`, title };
  if (cmp.stale) return { cls: "mx-stale", main, sub: "stale — refresh", title };
  const findings = rd + miss;
  const cls = findings === 0 ? "mx-match" : findings >= MX_HI ? "mx-diff-hi" : "mx-diff-lo";
  return { cls, main: findings === 0 ? "✓ 0 report ≠" : main, sub: parts.join(" · "), title };
}

// The "three sides" card: per report, its TSN dataset and its ArcGIS build.
function renderThreeWaySides(snap) {
  const lib = snap.library || {}, builds = lib.builds || {};
  const fams = [...new Set((snap.all_rows || []).map((r) => r.family))];
  const meta = snap.tsn_meta || {};
  const chips = $("twSideChips");
  chips.textContent = "";
  let ready = 0;
  fams.forEach((f) => {
    const b = builds[f] || {};
    const row = (snap.all_rows || []).find((r) => r.family === f) || {};
    const tsn = Object.entries(meta).find(([k]) => (snap.row_family || {})[k] === f);
    const tsnReady = !!(tsn && tsn[1].ready);
    const line = agBuildLine(b);
    const ok = tsnReady && b.built && !b.stale;
    if (ok) ready += 1;
    const chip = document.createElement("button");
    chip.className = "agb-chip" + (ok ? " ok" : " warn");
    const label = (row.label || f).replace(/ \(PDF\)$/, "");
    chip.textContent = `${label} — TSN ${tsnReady ? "✓" : "missing"} · ArcGIS ${line.text}`;
    chip.title = (tsn && tsn[1].source_path ? `TSN: ${tsn[1].source_path}\n` : "TSN: not in the library\n")
      + (b.path ? `ArcGIS: ${b.path}\n` : "") + "Open ArcGIS reports";
    chip.onclick = () => setArcgisSub("builds");
    chips.appendChild(chip);
  });
  $("twSidesMeta").textContent = `${ready} of ${fams.length} report${fams.length === 1 ? "" : "s"} ready`;
  $("twSidesHint").textContent = ready === fams.length
    ? "Each cell compares the day's export with the report's TSN dataset and its ArcGIS "
      + "build. 'report ≠' counts the cells where TSN and the layers agree and the report "
      + "doesn't — the ones worth chasing. Read the build's as-of date beside the export day."
    : "A report needs its TSN extract in the library (Settings ▸ TSN reports) and its ArcGIS "
      + "build (ArcGIS reports) before its cells can compare.";
}

async function renderThreeWayMatrix() {
  const grid = $("twMatrixGrid");
  if (!grid) return;
  const seq = ++_twRenderSeq;
  let snap;
  try { snap = await api.three_way_matrix_info(); } catch (e) {
    $("twSidesHint").textContent = "Status unavailable: " + e;
    return;
  }
  if (seq !== _twRenderSeq || !snap) return;     // a newer render started
  if (snap.error) { $("twSidesHint").textContent = snap.error; return; }
  TWM = snap;
  const days = snap.days || [], locked = !!(S.st && S.st.task);
  const builds = (snap.library || {}).builds || {};
  renderThreeWaySides(snap);

  const srcSel = $("twMatrixSource");
  srcSel.textContent = "";
  (snap.sources || []).forEach((s) => {
    const o = document.createElement("option");
    o.value = s.key; o.textContent = s.label;
    if (s.key === snap.source) o.selected = true;
    srcSel.appendChild(o);
  });
  srcSel.disabled = locked;
  srcSel.onchange = async () => {
    const r = await api.set_three_way_matrix_source(srcSel.value);
    if (r && r.error) showMessage("error", "Can't set source", r.error);
    await renderThreeWayMatrix();
  };

  const addSel = $("twMatrixAddDay"), addBtn = $("btnTwAddDay");
  const avail = (snap.available_days || []).filter((d) => !days.includes(d));
  addSel.textContent = "";
  if (!avail.length) {
    const o = document.createElement("option");
    o.value = ""; o.textContent = days.length ? "— no more exported days —" : "— no exported days —";
    addSel.appendChild(o);
  } else {
    avail.forEach((d) => {
      const o = document.createElement("option"); o.value = d; o.textContent = mxDayOptionText(d, snap);
      addSel.appendChild(o);
    });
  }
  addSel.disabled = locked || !avail.length;
  addBtn.disabled = locked || !avail.length;
  addBtn.onclick = () => mxBusyClick(addBtn, async () => {
    const d = addSel.value;
    if (!d) return;
    const r = await api.add_three_way_matrix_day(d);
    if (r && r.error) showMessage("error", "Can't add day", r.error);
    await renderThreeWayMatrix();
  });

  mxReportToggles("twMatrixReportToggles", "three_way", snap, locked,
    (key, next) => api.set_three_way_matrix_report(key, next), renderThreeWayMatrix);

  grid.textContent = "";
  if (!snap.rows.length) {
    mxNoRows(grid);
    wireThreeWayFooter();
    updateThreeWayMatrixProgress();
    return;
  }
  if (!days.length) {
    grid.style.gridTemplateColumns = ""; grid.style.gridTemplateRows = "";
    const empty = document.createElement("div");
    empty.className = "dm-empty";
    empty.textContent = "Add an export day from Matrix options to compare each Clean Road "
      + "export — Excel and PDF — with its TSN extract and its ArcGIS build at once.";
    grid.appendChild(empty);
    wireThreeWayFooter();
    updateThreeWayMatrixProgress();
    return;
  }
  grid.style.gridTemplateColumns = `minmax(230px,1.3fr) repeat(${days.length}, minmax(130px,1fr))`;
  grid.style.gridTemplateRows = ["auto"].concat(snap.rows.map(() => "minmax(54px,1fr)")).join(" ");

  const corner = document.createElement("div");
  corner.className = "mx-cell mx-corner mx-colhead";
  corner.textContent = "Report \\ Day";
  grid.appendChild(corner);
  days.forEach((d) => {
    const h = document.createElement("div");
    h.className = "mx-cell mx-colhead";
    const lab = document.createElement("div"); lab.className = "dnd-handle";
    lab.textContent = d;
    h.appendChild(lab);
    const siteTag = mxDaySiteTag(d, snap);
    if (siteTag) h.appendChild(siteTag);
    const btns = document.createElement("span"); btns.className = "mxch-btns";
    btns.append(
      mxHeadBtn("i-compare", `Rebuild every report for ${d} (three-way)`, "mxch-rebuild",
        async () => {
          const r = await api.rebuild_three_way_matrix("all", null, d);
          if (r && r.nothing) showMessage("info", "Nothing to rebuild", "No cell of this day has all three sides yet.");
          else if (r && r.error) showMessage("error", "Can't rebuild", r.error);
        }),
      mxRemoveColumnBtn(d, () => api.remove_three_way_matrix_day(d), renderThreeWayMatrix));
    h.appendChild(btns);
    dndAttach(h, h, d, "tw-day", "x", () => days.slice(), async (order) => {
      const r = await api.set_three_way_matrix_day_order(order);
      if (r && r.error) showMessage("error", "Can't reorder", r.error);
      else await renderThreeWayMatrix();
    });
    grid.appendChild(h);
  });

  const families = snap.row_family || {};
  snap.rows.forEach((rk) => {
    const rlabel = snap.row_labels[rk] || rk;
    const b = builds[families[rk] || rk] || {};
    const tsn = (snap.tsn_meta || {})[rk] || {};
    const rh = document.createElement("div");
    rh.className = "mx-cell mx-rowhead";
    rh.dataset.rk = rk; rh.dataset.label = rlabel;
    const top = document.createElement("div"); top.className = "mxrh-top";
    const lbl = document.createElement("span"); lbl.className = "mxrh-label";
    lbl.textContent = rlabel;
    top.appendChild(lbl);
    top.appendChild(mxHeadBtn("i-compare", `Rebuild ${rlabel} for every day (three-way)`,
      "mxch-rebuild", async () => {
        const r = await api.rebuild_three_way_matrix("all", rk, null);
        if (r && r.nothing) showMessage("info", "Nothing to rebuild", "No cell of this row has all three sides yet.");
        else if (r && r.error) showMessage("error", "Can't rebuild", r.error);
      }));
    rh.appendChild(top);
    const side = agBuildLine(b);
    const bl = document.createElement("div");
    const warn = !tsn.ready || !b.built || b.stale;
    bl.className = "mxrh-build" + (warn ? " warn" : " ok");
    const bt = document.createElement("span"); bt.className = "mxrh-buildtext";
    bt.textContent = `TSN ${tsn.ready ? "✓" : "missing"} · ArcGIS ${b.asof ? "as of " + b.asof : side.text}`;
    bt.title = (tsn.source_path ? `TSN: ${tsn.source_path}\n` : "TSN: not in the library\n")
      + `ArcGIS: ${side.text}` + (b.path ? `\n${b.path}` : "");
    bl.appendChild(bt);
    if (b.built) {
      bl.appendChild(mxHeadBtn("i-external", `Open the ${b.label || rlabel} ArcGIS build`, "mxch-open",
        async () => {
          const r = await api.open_arcgis_report(rk);
          if (r && r.error) showMessage("error", "Can't open", r.error);
        }));
    }
    rh.appendChild(bl);
    dndAttach(rh, top, rk, "tw-row", "y", () => snap.rows.slice(), async (order) => {
      const r = await api.set_three_way_matrix_row_order(order);
      if (r && r.error) showMessage("error", "Can't reorder", r.error);
      else await renderThreeWayMatrix();
    });
    grid.appendChild(rh);
    days.forEach((d) => {
      const c = snap.cells[rk][d], cmp = c.cmp || {};
      const cell = document.createElement("div"); cell.className = "mx-cell";
      const main = document.createElement("div"); main.className = "mx-num";
      const sub = document.createElement("div"); sub.className = "mx-sub";
      const v = twCellContent(cmp);
      cell.classList.add(v.cls); main.textContent = v.main; sub.textContent = v.sub;
      if (v.warn) cell.classList.add(v.warn);
      const ed = c.export || {};
      cell.title = `${rlabel} — ${d}: TSMIS vs TSN vs ArcGIS\nExport: `
        + (ed.present ? fmtAge(ed.age_seconds) : "not exported")
        + (b.built ? `\nArcGIS build: ${side.text}` : "")
        + (v.title ? `\n\n${v.title}` : "");
      cell.append(main, sub);
      const acts = document.createElement("div"); acts.className = "mx-actions";
      if (!cmp.missing_side) {
        acts.appendChild(mxActBtn("i-compare", "Build / rebuild this three-way comparison", false,
          async () => {
            const r = await api.build_three_way_matrix_cell(rk, d);
            if (r && r.error) showMessage("error", "Can't build", r.error);
          }));
      }
      if (cmp.built) {
        const ob = mxActBtn("i-external", "Open this three-way workbook", false, async () => {
          const r = await api.open_three_way_cell_comparison(rk, d);
          if (r && r.error) showMessage("error", "Can't open", r.error);
        });
        ob.classList.add("mx-open"); acts.appendChild(ob);
      }
      cell.appendChild(acts);
      grid.appendChild(cell);
    });
  });

  wireThreeWayFooter();
  updateThreeWayMatrixProgress();
}

function wireThreeWayFooter() {
  $("btnTwMxBuildAll").onclick = async () => {
    const r = await api.rebuild_three_way_matrix("all");
    if (r && r.nothing) showMessage("info", "Nothing to build", "Add an exported day, and make sure each report has its TSN extract and its ArcGIS build.");
    else if (r && r.error) showMessage("error", "Can't build", r.error);
  };
  $("btnTwMxRebuildAll").onclick = async () => {
    const r = await api.rebuild_three_way_matrix("stale");
    if (r && r.nothing) showMessage("info", "Up to date", "Every three-way comparison is current.");
    else if (r && r.error) showMessage("error", "Can't rebuild", r.error);
  };
  $("btnTwMxOpenComparisons").onclick = async () => {
    const r = await api.open_three_way_comparisons_folder();
    if (r && r.error) showMessage("error", "Can't open", r.error);
  };
  $("btnTwMxCancel").onclick = () => api.cancel_run();
}

// Called on every state push while the sub-tab shows (no API call): the live
// progress line, the Cancel button, the queue panel and the lock-sensitive bits.
function updateThreeWayMatrixProgress() {
  const el = $("twMatrixProgress");
  if (el) {
    const m = S.st && S.st.matrix;
    if (m && m.total) {
      el.hidden = false;
      el.textContent = m.phase === "building" ? agbProgressText(m)
        : (m.phase === "layers" || m.phase === "probe") ? aglProgressText(m)
        : mxProgressText(m);
    } else el.hidden = true;
  }
  const locked = !!(S.st && S.st.task);
  document.querySelectorAll(
    "#twMatrixSource, #twMatrixReportToggles .mx-toggle, #twMatrixGrid .mxrh-top .mxch-rebuild")
    .forEach((c) => { c.disabled = locked; });
  const addSel = $("twMatrixAddDay"), addBtn = $("btnTwAddDay");
  const noAvail = !addSel || !addSel.querySelector('option[value]:not([value=""])');
  if (addSel) addSel.disabled = locked || noAvail;
  if (addBtn) addBtn.disabled = locked || noAvail;
  const cancel = $("btnTwMxCancel");
  if (cancel) {
    const running = !!(S.st && S.st.task === "matrix");
    cancel.classList.toggle("hidden", !running);
    cancel.disabled = !running;
  }
  renderQueuePanel("twQueueGroup", "twQueue", "twQueueCount");
}

function bindThreeWay() {
  $("subAgThreeWay").onclick = () => setArcgisSub("threeway");
  $("btnTwGoBuilds").onclick = () => setArcgisSub("builds");
}
