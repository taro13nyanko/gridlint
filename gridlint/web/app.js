/* Gridlint front end. Vanilla ES modules-free JS: no build step, no framework. */
'use strict';

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

const state = {
  report: null,
  selected: null,
  sheet: null,
  severities: new Set(['critical', 'warning', 'info']),
  showFormulas: false,
  signedIn: false,
};

/* ------------------------------------------------------------------ helpers */

function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function num(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v !== 'number') return String(v);
  const abs = Math.abs(v);
  if (Number.isInteger(v)) return v.toLocaleString('en-US');
  if (abs > 0 && abs < 0.001) return v.toExponential(2);
  return v.toLocaleString('en-US', { maximumFractionDigits: abs < 10 ? 3 : 1 });
}

function money(v, isCurrency) {
  if (v === null || v === undefined) return '';
  const s = Math.round(Math.abs(v)).toLocaleString('en-US');
  return isCurrency ? `$${s}` : s;
}

/* Show a value the way the cell shows it: 0.063 as 6.3%, 243659 as $243,659. */
function fmtValue(v, kind) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v !== 'number') return String(v);
  if (kind === 'percent') {
    const p = v * 100;
    return `${p.toLocaleString('en-US', { maximumFractionDigits: Math.abs(p) < 10 ? 1 : 0 })}%`;
  }
  if (kind === 'currency') {
    return (v < 0 ? '-$' : '$') + Math.round(Math.abs(v)).toLocaleString('en-US');
  }
  return num(v);
}

function toast(msg, bad = false) {
  const el = $('#toast');
  el.textContent = msg;
  el.classList.toggle('bad', bad);
  el.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { el.hidden = true; }, 4200);
}

function loading(on, text) {
  $('#loading').hidden = !on;
  if (text) $('#loading-text').textContent = text;
}

async function api(path, opts = {}) {
  const res = await fetch(path, { credentials: 'same-origin', ...opts });
  let body = null;
  try { body = await res.json(); } catch { /* not json */ }
  if (!res.ok) throw new Error(body?.detail || `${res.status} ${res.statusText}`);
  return body;
}

/* -------------------------------------------------------------------- views */

function showLanding() {
  $('#landing').hidden = false;
  $('#report').hidden = true;
  window.scrollTo({ top: 0, behavior: 'instant' });
}

function showReport(report) {
  state.report = report;
  state.sheet = Object.keys(report.sheets || {})[0] || null;
  state.selected = report.findings[0] || null;
  $('#landing').hidden = true;
  $('#report').hidden = false;
  renderReport();
  window.scrollTo({ top: 0, behavior: 'instant' });
}

/* ------------------------------------------------------------------ rendering */

function renderReport() {
  const r = state.report;
  $('#rep-name').textContent = r.path;
  $('#rep-sub').textContent =
    `${r.stats.formulas} formulas across ${r.stats.sheets} sheet${r.stats.sheets === 1 ? '' : 's'}` +
    ` · checked in ${r.duration_ms} ms`;

  const eng = r.engine;
  const badges = [];
  if (eng.checked > 0) {
    badges.push(`<span class="badge ${eng.trustworthy ? 'ok' : 'warn'}" title="Our engine recomputed every formula and compared it with the values the spreadsheet app saved in the file.">
      ${eng.trustworthy ? '✓' : '!'} engine matches the file on ${eng.matched}/${eng.checked} values</span>`);
  } else {
    badges.push(`<span class="badge warn" title="This file has no saved results to compare against, so the engine could not be cross-checked.">no saved values to check against</span>`);
  }
  if (r.counts.critical) badges.push(`<span class="badge crit">${r.counts.critical} critical</span>`);
  if (r.counts.warning) badges.push(`<span class="badge warn">${r.counts.warning} warning</span>`);
  if (r.counts.info) badges.push(`<span class="badge">${r.counts.info} info</span>`);
  if (!r.findings.length) badges.push('<span class="badge ok">✓ no defects</span>');
  $('#rep-badges').innerHTML = badges.join('');

  renderHeadline();
  renderFilters();
  renderFindingList();
  renderDetail();
  renderSheetTabs();
  renderGrid();
}

function renderHeadline() {
  const r = state.report;
  const top = r.findings.find(f => f.headline_changes?.length) || r.findings[0];
  const box = $('#rep-headline');
  if (!top || !top.headline_changes?.length) { box.hidden = true; return; }
  box.hidden = false;
  $('#headline-text').textContent = top.detail.split('.')[0] + '.';
  $('#headline-changes').innerHTML = top.headline_changes.slice(0, 4).map(h =>
    `<li><b>${esc(h.label)}</b> <span class="from">${esc(fmtValue(h.before, h.fmt))}</span> → <span class="to">${esc(fmtValue(h.after, h.fmt))}</span></li>`
  ).join('');
}

function renderFilters() {
  const counts = state.report.counts;
  const defs = [['critical', counts.critical], ['warning', counts.warning], ['info', counts.info]];
  $('#sev-filters').innerHTML = defs.map(([sev, n]) =>
    `<button class="chip" data-sev="${sev}" aria-pressed="${state.severities.has(sev)}">${sev} ${n}</button>`
  ).join('');
  $$('#sev-filters .chip').forEach(btn => btn.addEventListener('click', () => {
    const sev = btn.dataset.sev;
    state.severities.has(sev) ? state.severities.delete(sev) : state.severities.add(sev);
    btn.setAttribute('aria-pressed', state.severities.has(sev));
    renderFindingList();
  }));
}

function visibleFindings() {
  return state.report.findings.filter(f => state.severities.has(f.severity));
}

function renderFindingList() {
  const list = visibleFindings();
  $('#no-findings').hidden = list.length > 0;
  $('#no-findings').textContent = state.report.findings.length
    ? 'Nothing at the selected severity.' : 'No defects found in this workbook.';
  $('#finding-list').innerHTML = list.map(f => {
    const meta = [];
    if (f.group_size > 1) meta.push(`${f.group_size} cells`);
    if (f.impact_cells) meta.push(`${f.impact_cells} cells change`);
    meta.push(f.rule);
    const impact = f.impact_value
      ? `<span class="f-money">${money(f.impact_value, f.impact_currency)}</span>` : '';
    return `<li><button class="f-btn" data-id="${esc(f.id)}" aria-current="${state.selected?.id === f.id}">
      <span class="f-top"><span class="dot ${f.severity}"></span><span class="f-title">${esc(f.title)}</span></span>
      <span class="f-meta"><span class="mono">${esc(f.cell)}</span>${impact}<span>${meta.join(' · ')}</span></span>
    </button></li>`;
  }).join('');
  $$('#finding-list .f-btn').forEach(btn => btn.addEventListener('click', () => {
    state.selected = state.report.findings.find(f => f.id === btn.dataset.id);
    const sheet = state.selected.sheet;
    if (sheet && state.sheet !== sheet) { state.sheet = sheet; renderSheetTabs(); }
    renderFindingList();
    renderDetail();
    renderGrid();
    scrollToFlagged();
  }));
}

function renderDetail() {
  const f = state.selected;
  const el = $('#detail');
  if (!f) {
    el.innerHTML = `<p class="empty">Select a finding to see what it changes.</p>`;
    return;
  }
  const parts = [];
  parts.push(`<h3>${esc(f.title)}</h3>
    <p class="where">${esc(f.cell)}${f.group_size > 1 ? ` and ${f.group_size - 1} more` : ''} · rule ${esc(f.rule)} · confidence ${Math.round(f.confidence * 100)}%</p>`);

  parts.push(`<div class="detail-sec"><p>${esc(f.detail)}</p></div>`);

  if (f.explanation) {
    parts.push(`<div class="detail-sec"><div class="note"><span class="who">Plain-English note</span>${esc(f.explanation)}</div></div>`);
  } else if (f.explanation_rejected?.length) {
    parts.push(`<div class="detail-sec"><div class="note rejected"><span class="who">Note discarded</span>
      The model's sentence quoted ${f.explanation_rejected.map(esc).join(', ')}, which is not in the measured
      evidence, so it was thrown away rather than shown to you.</div></div>`);
  }

  if (f.formula) {
    parts.push(`<div class="detail-sec"><h4>Formula</h4><div class="formula">${esc(f.formula)}</div>`);
    if (f.fix) {
      parts.push(`<p class="arrow">↓ ${esc(f.fix.label)}</p>
        <div class="formula after">${esc(f.fix.new_formula)}</div>`);
    }
    parts.push('</div>');
  }

  if (f.headline_changes?.length) {
    parts.push(`<div class="detail-sec"><h4>Numbers people read that change</h4>
      <table class="diff"><thead><tr><th>Line</th><th>Now</th><th>After the fix</th></tr></thead><tbody>
      ${f.headline_changes.map(h => `<tr><td>${esc(h.label)}</td><td class="from">${esc(fmtValue(h.before, h.fmt))}</td><td class="to">${esc(fmtValue(h.after, h.fmt))}</td></tr>`).join('')}
      </tbody></table></div>`);
  }

  if (f.fix) {
    const v = f.fix_verified;
    parts.push(`<div class="detail-sec">
      <span class="verified ${v ? '' : 'unverified'}">${v ? '✓ fix verified by recalculation' : '! fix not verified'}</span>
      <p class="fine" style="margin-top:8px">${v
        ? `The fix was applied to a copy, the workbook recomputed, and ${f.impact_cells} cell${f.impact_cells === 1 ? '' : 's'} changed. No new errors appeared.`
        : 'This suggestion did not survive recalculation, so it is shown but not offered for download.'}</p>
    </div>`);
  }

  if (f.trace?.length) {
    const head = f.headline_changes?.[0];
    parts.push(`<div class="detail-sec"><h4>Where ${esc(head ? head.label : 'that number')} comes from</h4>
      <ol class="trace">${f.trace.map(s => `
        <li class="${s.changed ? 'moves' : ''}" style="--d:${Math.min(s.depth, 5)}">
          <span class="t-cell mono">${esc(s.cell.split('!').pop())}</span>
          <span class="t-label">${esc(s.label || (s.is_input ? 'input' : ''))}</span>
          <span class="t-formula mono">${esc(s.formula || num(s.value))}</span>
        </li>`).join('')}</ol>
      <p class="fine">Highlighted rows are the ones whose value moves when this fix is applied.</p>
    </div>`);
  }

  if (f.fix_diff?.length) {
    parts.push(`<details class="detail-sec"><summary class="fine">Every cell that changes (${f.impact_cells})</summary>
      <table class="diff"><tbody>${f.fix_diff.map(d =>
        `<tr><td class="cell">${esc(d.cell)}</td><td class="from">${esc(fmtValue(d.before, d.fmt))}</td><td class="to">${esc(fmtValue(d.after, d.fmt))}</td></tr>`).join('')}
      </tbody></table></details>`);
  }

  parts.push(`<div class="detail-actions">
    <button class="btn sm" id="btn-explain">Explain in plain English</button>
    ${state.signedIn && f.fix && f.fix_verified ? '<button class="btn sm" id="btn-download">Download corrected file</button>' : ''}
  </div>`);

  el.innerHTML = parts.join('');
  $('#btn-explain')?.addEventListener('click', explainAll);
  $('#btn-download')?.addEventListener('click', downloadFixed);
}

function renderSheetTabs() {
  const sheets = Object.keys(state.report.sheets || {});
  $('#sheet-tabs').innerHTML = sheets.map(s =>
    `<button class="chip" data-sheet="${esc(s)}" aria-pressed="${state.sheet === s}">${esc(s)}</button>`).join('');
  $$('#sheet-tabs .chip').forEach(b => b.addEventListener('click', () => {
    state.sheet = b.dataset.sheet;
    renderSheetTabs();
    renderGrid();
  }));
}

function colName(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function renderGrid() {
  const sheet = state.report.sheets?.[state.sheet];
  const table = $('#grid');
  if (!sheet) { table.innerHTML = ''; return; }
  const flagged = new Set();
  for (const f of state.report.findings) {
    for (const c of (f.group_cells?.length ? f.group_cells : [f.cell])) flagged.add(c);
  }
  const selectedCells = new Set(state.selected
    ? (state.selected.group_cells?.length ? state.selected.group_cells : [state.selected.cell]) : []);

  const head = ['<tr><th class="rownum"></th>'];
  for (let c = 1; c <= sheet.shown_cols; c++) {
    head.push(`<th class="${c === 1 ? 'labelcol' : ''}">${colName(c)}</th>`);
  }
  head.push('</tr>');

  let firstSelected = true;
  const rows = sheet.rows.map((row, i) => {
    const cells = row.map((cell, j) => {
      if (!cell) return `<td class="${j === 0 ? 'labelcol' : ''}"></td>`;
      const ref = `${state.sheet}!${colName(j + 1)}${i + 1}`;
      const isFlag = flagged.has(ref);
      const isSel = selectedCells.has(ref);
      const cls = [
        j === 0 ? 'labelcol' : '',
        typeof cell.v === 'string' ? 'text' : '',
        cell.f ? 'f' : '',
        isFlag ? 'flag' : '',
        state.showFormulas && cell.f ? 'formula-view' : '',
      ].filter(Boolean).join(' ');
      const shown = state.showFormulas && cell.f ? cell.f : num(cell.v);
      const title = cell.f ? ` title="${esc(cell.f)}"` : '';
      // Only one element may carry the id the scroll target uses.
      const id = isSel && firstSelected ? (firstSelected = false, ' id="flagged-cell"') : '';
      return `<td class="${cls}"${title}${id}>${esc(shown)}</td>`;
    });
    return `<tr><td class="rownum">${i + 1}</td>${cells.join('')}</tr>`;
  });

  table.innerHTML = `<thead>${head.join('')}</thead><tbody>${rows.join('')}</tbody>`;
  if (sheet.max_row > sheet.shown_rows || sheet.max_col > sheet.shown_cols) {
    const note = document.createElement('caption');
    note.className = 'fine';
    note.style.cssText = 'caption-side:bottom;text-align:left;padding:8px 10px';
    note.textContent = `Showing the first ${sheet.shown_rows} rows and ${sheet.shown_cols} columns of ${sheet.max_row} × ${sheet.max_col}. Every cell was checked.`;
    table.appendChild(note);
  }
}

function scrollToFlagged() {
  const el = $('#flagged-cell');
  if (el) el.scrollIntoView({ block: 'center', inline: 'center', behavior: 'smooth' });
}

/* -------------------------------------------------------------------- actions */

async function runDemo() {
  loading(true, 'Reading formulas and recalculating…');
  try {
    showReport(await api('/api/demo'));
  } catch (e) {
    toast(e.message, true);
  } finally {
    loading(false);
  }
}

async function checkFile(file) {
  if (!file) return;
  if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
    toast('Gridlint reads .xlsx and .xlsm files. Save as Excel Workbook and try again.', true);
    return;
  }
  loading(true, `Checking ${file.name}…`);
  const fd = new FormData();
  fd.append('file', file);
  try {
    const path = state.signedIn ? '/api/workbooks' : '/api/check';
    const res = await api(path, { method: 'POST', body: fd });
    showReport(res.report || res);
    if (res.workbook_id) state.workbookId = res.workbook_id;
  } catch (e) {
    toast(e.message, true);
  } finally {
    loading(false);
  }
}

async function explainAll() {
  loading(true, 'Writing the plain-English notes…');
  try {
    const res = await api('/api/explain', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        findings: state.report.findings.slice(0, 8),
        report: state.report,
        context: `Workbook ${state.report.path}`,
      }),
    });
    const byCell = new Map(res.explanations.map(e => [`${e.rule}:${e.cell}`, e]));
    let written = 0, rejected = 0;
    for (const f of state.report.findings) {
      const e = byCell.get(`${f.rule}:${f.cell}`);
      if (!e) continue;
      f.explanation = e.explanation;
      f.explanation_rejected = e.rejected_numbers;
      if (e.explanation) written++;
      if (e.rejected_numbers?.length) rejected++;
    }
    renderDetail();
    if (!written && !rejected) {
      toast('No model is configured, so the built-in explanation is shown instead.');
    } else {
      toast(`${written} note${written === 1 ? '' : 's'} written` +
        (rejected ? `, ${rejected} discarded for quoting an unverified number` : ''));
    }
  } catch (e) {
    toast(e.message, true);
  } finally {
    loading(false);
  }
}

async function downloadFixed() {
  if (!state.workbookId) { toast('Sign in and upload the file to download a corrected copy.', true); return; }
  loading(true, 'Applying the verified fixes…');
  try {
    const res = await fetch(`/api/workbooks/${state.workbookId}/fixed`, {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ finding_ids: [state.selected.id] }),
    });
    if (!res.ok) throw new Error((await res.json()).detail || 'Could not build the corrected file.');
    const blob = await res.blob();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = state.report.path.replace(/\.xlsx?$/i, '') + '-fixed.xlsx';
    a.click();
    URL.revokeObjectURL(a.href);
    toast('Downloaded. Open it in Excel to recalculate.');
  } catch (e) {
    toast(e.message, true);
  } finally {
    loading(false);
  }
}

/* ----------------------------------------------------------------------- auth */

let authMode = 'login';

function openAuth(mode = 'login') {
  authMode = mode;
  $('#auth-title').textContent = mode === 'login' ? 'Sign in' : 'Create a workspace';
  $('#auth-submit').textContent = mode === 'login' ? 'Sign in' : 'Create workspace';
  $('#auth-switch').textContent = mode === 'login' ? 'Create a workspace instead' : 'I already have an account';
  $('#ws-field').hidden = mode === 'login';
  $('#auth-error').hidden = true;
  $('#auth-dialog').showModal();
}

async function submitAuth() {
  const form = $('#auth-form');
  const fd = new FormData(form);
  if (authMode === 'login') fd.delete('workspace');
  try {
    await api(authMode === 'login' ? '/api/login' : '/api/signup', { method: 'POST', body: fd });
    $('#auth-dialog').close();
    await refreshAccount();
    toast('Signed in. Uploads are now saved to your workspace.');
  } catch (e) {
    const err = $('#auth-error');
    err.textContent = e.message;
    err.hidden = false;
  }
}

async function refreshAccount() {
  try {
    const me = await api('/api/me');
    if (!me.signed_in) throw new Error('anonymous');
    state.signedIn = true;
    $('#nav-account').textContent = 'Sign out';
    $('#nav-account').onclick = async () => {
      await api('/api/logout', { method: 'POST' });
      state.signedIn = false;
      $('#nav-account').textContent = 'Sign in';
      $('#nav-account').onclick = () => openAuth('login');
      toast('Signed out.');
    };
    return me;
  } catch {
    state.signedIn = false;
    $('#nav-account').textContent = 'Sign in';
    $('#nav-account').onclick = () => openAuth('login');
    return null;
  }
}

/* ----------------------------------------------------------------------- init */

function init() {
  $('#try-demo').addEventListener('click', runDemo);
  $('#pick-file').addEventListener('click', () => $('#file-input').click());
  $('#back-home').addEventListener('click', showLanding);
  $('#show-formulas').addEventListener('change', e => {
    state.showFormulas = e.target.checked;
    renderGrid();
  });

  const dz = $('#dropzone');
  dz.addEventListener('click', () => $('#file-input').click());
  dz.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('#file-input').click(); }
  });
  ['dragenter', 'dragover'].forEach(t => dz.addEventListener(t, e => {
    e.preventDefault(); dz.classList.add('over');
  }));
  ['dragleave', 'drop'].forEach(t => dz.addEventListener(t, e => {
    e.preventDefault(); dz.classList.remove('over');
  }));
  dz.addEventListener('drop', e => checkFile(e.dataTransfer.files[0]));
  $('#file-input').addEventListener('change', e => checkFile(e.target.files[0]));

  $('#auth-cancel').addEventListener('click', () => $('#auth-dialog').close());
  $('#auth-submit').addEventListener('click', submitAuth);
  $('#auth-form').addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); submitAuth(); }
  });
  $('#auth-switch').addEventListener('click', e => {
    e.preventDefault();
    $('#auth-dialog').close();
    openAuth(authMode === 'login' ? 'signup' : 'login');
  });

  refreshAccount();

  const shared = location.pathname.match(/^\/report\/(.+)$/);
  if (shared) {
    loading(true, 'Opening the shared report…');
    api(`/api/shared/${shared[1]}`)
      .then(r => showReport(r.report))
      .catch(e => toast(e.message, true))
      .finally(() => loading(false));
  }
}

document.addEventListener('DOMContentLoaded', init);
