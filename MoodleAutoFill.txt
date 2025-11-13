// ==UserScript==
// @name         Moodle Grader (Single Long Page) – Sequential CSV Autofill
// @namespace    lpubatangas.moodle.autofill.singlepage
// @version      1.6
// @description  Fill mark + comment for stacked single-page manual grading from a local CSV, in order (no email matching).
// @match        https://lms.lpubatangas.edu.ph/mod/quiz/report.php*mode=grading*
// @grant        none
// @require      https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js
// ==/UserScript==

(function () {
  'use strict';

  // --- Quick sanity: confirm we're on the stacked-single page (we expect multiple -mark inputs)
  const markInputs = Array.from(document.querySelectorAll('input[id$="-mark"]'));
  const commentAreas = Array.from(document.querySelectorAll('textarea[id$="_-comment_id"]'));

  if (markInputs.length === 0 || commentAreas.length === 0) {
    // Not the stacked single page (or Moodle theme differs)
    // We still install the panel so user sees the hint.
    console.warn('[Moodle Autofill] Expected grade/comment fields not found on this page.');
  }

  // --- Floating control panel
  const panel = document.createElement('div');
  panel.style.cssText = `
    position: fixed; z-index: 2147483647; right: 16px; bottom: 16px;
    background: #111827; color: #fff; padding: 12px; border-radius: 12px;
    box-shadow: 0 8px 20px rgba(0,0,0,.35); width: 320px; font: 13px/1.3 system-ui, -apple-system, Segoe UI, Roboto, Arial;
  `;
  panel.innerHTML = `
    <div style="font-weight:700;margin-bottom:6px;">Moodle Autofill (CSV → Page)</div>

    <label style="display:block;margin-bottom:6px;">
      <span style="opacity:.85;">CSV file</span><br>
      <input type="file" id="mf-csv" accept=".csv" style="width:100%; margin-top:4px;">
    </label>

    <div style="display:flex; gap:8px; align-items:center; margin-bottom:8px;">
      <label style="flex:1;">
        <span style="opacity:.85;">Row offset</span><br>
        <input id="mf-offset" type="number" min="0" value="0" style="width:100%; margin-top:4px; padding:4px 6px; border-radius:6px; border:1px solid #374151; background:#0b1220; color:#fff;">
      </label>
      <label style="flex:1;">
        <span style="opacity:.85;">Max rows</span><br>
        <input id="mf-max" type="number" min="1" value="25" style="width:100%; margin-top:4px; padding:4px 6px; border-radius:6px; border:1px solid #374151; background:#0b1220; color:#fff;">
      </label>
    </div>

    <div style="display:flex; gap:8px; margin-bottom:8px;">
      <button id="mf-fill" disabled style="flex:1;background:#2563eb;border:none;color:#fff;padding:8px;border-radius:8px;cursor:pointer;">Fill</button>
      <button id="mf-clear" style="background:#374151;border:none;color:#fff;padding:8px;border-radius:8px;cursor:pointer;">Clear</button>
    </div>

    <div id="mf-status" style="font-size:12px; opacity:.9;">No CSV loaded.</div>
  `;
  document.body.appendChild(panel);

  const fileInput  = panel.querySelector('#mf-csv');
  const fillBtn    = panel.querySelector('#mf-fill');
  const clearBtn   = panel.querySelector('#mf-clear');
  const statusEl   = panel.querySelector('#mf-status');
  const offsetEl   = panel.querySelector('#mf-offset');
  const maxEl      = panel.querySelector('#mf-max');

  let csvRows = [];  // parsed rows, in order

  function setStatus(msg) { statusEl.textContent = msg; }
  function fire(el, type)  { el && el.dispatchEvent(new Event(type, { bubbles: true })); }

  // Convert plain → safe HTML for TinyMCE/Atto
  function toHTML(text) {
    if (text == null) return '';
    if (/[<>]/.test(text)) return text; // already looks like HTML
    const esc = text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    return '<p>' + esc.replace(/\r\n/g, '\n').split('\n\n').map(p => p.replace(/\n/g, '<br>')).join('</p><p>') + '</p>';
  }

  // TinyMCE iframe setter (visible editor)
  function setTinyMCEFromTextareaId(textareaId, html) {
    const ifr = document.getElementById(textareaId + '_ifr') || document.querySelector(`iframe#${CSS.escape(textareaId)}_ifr`);
    if (!ifr) return;
    try {
      const doc = ifr.contentDocument || ifr.contentWindow.document;
      if (doc && doc.body) {
        doc.body.innerHTML = html;
        // Fire a few events so the editor notices
        doc.body.dispatchEvent(new Event('input', { bubbles: true }));
        doc.body.dispatchEvent(new Event('keyup', { bubbles: true }));
        doc.body.dispatchEvent(new Event('change', { bubbles: true }));
      }
    } catch (e) {
      console.warn('[Moodle Autofill] Could not access TinyMCE iframe:', e);
    }
  }

  // --- CSV loading
  function normalizeHeader(h) {
    h = (h || '').trim().toLowerCase();
    if (['mark','score','grade','points'].includes(h)) return 'mark';
    if (['comment','feedback','remarks','note','notes'].includes(h)) return 'comment';
    if (['email','mail'].includes(h)) return 'email';
    return h;
  }

  fileInput.addEventListener('change', () => {
    const file = fileInput.files && fileInput.files[0];
    if (!file) { csvRows = []; fillBtn.disabled = true; setStatus('No CSV loaded.'); return; }

    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      transformHeader: normalizeHeader,
      complete: (res) => {
        if (res.errors?.length) {
          console.error(res.errors);
          setStatus('CSV parse error: ' + res.errors[0].message);
          fillBtn.disabled = true;
          return;
        }
        csvRows = res.data.map(r => ({
          email: (r.email ?? '').toString().trim(),
          mark:  (r.mark  ?? '').toString().trim(),
          comment: (r.comment ?? '').toString()
        }));
        if (!csvRows.length) {
          setStatus('CSV is empty.'); fillBtn.disabled = true; return;
        }
        fillBtn.disabled = false;
        setStatus(`CSV loaded: ${csvRows.length} rows. Found ${markInputs.length} mark fields & ${commentAreas.length} comment fields on page.`);
      }
    });
  });

  // --- Clear (panel state only)
  clearBtn.addEventListener('click', () => {
    fileInput.value = '';
    csvRows = [];
    fillBtn.disabled = true;
    setStatus('No CSV loaded.');
  });

  // --- Main fill
  fillBtn.addEventListener('click', () => {
    if (!csvRows.length) { setStatus('Load a CSV first.'); return; }

    const marks = Array.from(document.querySelectorAll('input[id$="-mark"]'));
    const comments = Array.from(document.querySelectorAll('textarea[id$="_-comment_id"]'));

    if (!marks.length || !comments.length || marks.length !== comments.length) {
      alert('Could not find the expected grade/comment fields for the stacked page. Check that all attempts are expanded/visible.');
      return;
    }

    const offset = Math.max(0, parseInt(offsetEl.value || '0', 10) || 0);
    const max    = Math.max(1, parseInt(maxEl.value || '25', 10) || 25);

    const fillCount = Math.min(max, marks.length, comments.length, Math.max(0, csvRows.length - offset));
    if (fillCount <= 0) { setStatus('Nothing to fill (check offset / CSV length).'); return; }

    let filled = 0;

    for (let i = 0; i < fillCount; i++) {
      const row = csvRows[offset + i];
      const markEl = marks[i];
      const taEl   = comments[i];

      // 1) set mark
      markEl.focus();
      markEl.value = row.mark ?? '';
      fire(markEl, 'input'); fire(markEl, 'change');

      // 2) set comment (textarea + TinyMCE iframe)
      const html = toHTML(row.comment ?? '');
      taEl.value = html;
      fire(taEl, 'input'); fire(taEl, 'change');
      setTinyMCEFromTextareaId(taEl.id, html);

      filled++;
    }

    setStatus(`Filled ${filled} row(s) starting at CSV row ${offset + 1}. Page shows ${marks.length} items.`);
    // Optional: scroll back to top for review
    // window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  // Heads-up if pagesize isn’t 25 (still okay, just FYI)
  try {
    const params = new URLSearchParams(location.search);
    const pg = params.get('pagesize');
    if (pg && pg !== '25') {
      setStatus((statusEl.textContent ? statusEl.textContent + ' · ' : '') + `Heads up: pagesize=${pg}. Script fills whatever is visible.`);
    }
  } catch {}
})();
