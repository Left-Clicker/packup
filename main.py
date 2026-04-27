#!/usr/bin/env python3
"""
Crowdin Translation Checker
Paste Chinese + English draft → 3-column table with glossary terms → copy back to Crowdin.
Glossary auto-loaded from "Mafia War's Glossary.csv" in the same directory.
"""

import csv
import http.server
import json
import os
import subprocess
import sys
import threading
import time
import webbrowser

PORT = 8765
# When packaged with PyInstaller (--onefile), __file__ points to a temporary
# extraction directory, not where the user actually placed the binary. Use the
# executable's directory in that case so the glossary CSV can sit next to the
# binary and be edited without rebuilding.
if getattr(sys, "frozen", False):
    SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
GLOSSARY_FILE = "Mafia War's Glossary.csv"


# Crowdin glossary export columns are fixed:
#   A (0)  -> Term [zh-CN]    源术语（中文）
#   M (12) -> Term [en-US]    目标术语（英文）
ZH_COL = 0
EN_COL = 12


def load_glossary():
    path = os.path.join(SCRIPT_DIR, GLOSSARY_FILE)
    if not os.path.exists(path):
        return [], f"Glossary file not found: {GLOSSARY_FILE}"

    with open(path, newline='', encoding='utf-8-sig') as f:
        rows = list(csv.reader(f))

    if len(rows) < 2:
        return [], "Glossary is empty or contains only a header row"

    terms = []
    for row in rows[1:]:
        zh = row[ZH_COL].strip() if len(row) > ZH_COL else ''
        en = row[EN_COL].strip() if len(row) > EN_COL else ''
        if zh and en:
            terms.append({'zh': zh, 'en': en})

    return terms, f"Loaded {len(terms)} glossary terms"


# ─────────────────────────────────────────────────────────────────────────────
# Embedded frontend
# ─────────────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Crowdin Translation Checker</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
     background:#0d1117;color:#c9d1d9;min-height:100vh}

/* ── Header ── */
.hdr{background:#161b22;border-bottom:1px solid #30363d;padding:12px 24px;
     display:flex;align-items:center;gap:16px;position:sticky;top:0;z-index:100}
.hdr h1{font-size:15px;font-weight:600;color:#f0f6fc;white-space:nowrap}
.gloss-tag{margin-left:auto;font-size:12px;padding:4px 12px;border-radius:20px;
           background:#122d20;color:#56d364;white-space:nowrap}
.gloss-tag.warn{background:#32100b;color:#f85149}
.gloss-tag.loading{background:#1c2a3a;color:#8b949e}

/* ── Buttons ── */
.btn{padding:7px 16px;border-radius:6px;border:none;cursor:pointer;
     font-size:13px;font-weight:500;transition:background .15s;white-space:nowrap}
.btn:disabled{opacity:.4;cursor:not-allowed}
.btn-blue{background:#1f6feb;color:#fff}.btn-blue:hover:not(:disabled){background:#388bfd}
.btn-green{background:#238636;color:#fff}.btn-green:hover:not(:disabled){background:#2ea043}
.btn-sky{background:#0969da;color:#fff}.btn-sky:hover:not(:disabled){background:#1f6feb}
.btn-gray{background:#21262d;color:#c9d1d9;border:1px solid #30363d}
.btn-gray:hover:not(:disabled){background:#30363d}
.btn-sm{padding:4px 10px;font-size:11px;border-radius:5px;border:none;cursor:pointer;
        font-weight:500;transition:background .15s;white-space:nowrap;
        background:#21262d;color:#8b949e;border:1px solid #30363d}
.btn-sm:hover{background:#30363d;color:#c9d1d9}
.btn-sm.active{background:#1c2a3a;color:#79c0ff;border-color:#1f4068}

/* ── Input panel ── */
.panel{padding:16px 24px;background:#161b22;border-bottom:1px solid #30363d}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:12px}
.field label{display:block;font-size:11px;color:#8b949e;margin-bottom:6px;
             font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.field label span{text-transform:none;font-weight:400;color:#484f58}
textarea{width:100%;background:#0d1117;border:1px solid #30363d;border-radius:6px;
         padding:9px 11px;color:#c9d1d9;font-size:12.5px;
         font-family:'SFMono-Regular',Consolas,monospace;resize:vertical;
         min-height:130px;line-height:1.6;overflow-y:auto;overflow-x:hidden}
textarea:focus{outline:none;border-color:#388bfd}
.row-btns{display:flex;gap:10px;align-items:center}
.status{font-size:13px;padding:5px 12px;border-radius:6px;display:none}
.status.info{background:#1c3557;color:#79c0ff;display:inline-block}
.status.ok  {background:#122d20;color:#56d364;display:inline-block}
.status.err {background:#32100b;color:#f85149;display:inline-block}

/* ── Blocks wrapper ── */
.blocks-wrap{padding:16px 24px}
.blocks-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}
.blocks-hdr>span{font-size:13px;color:#8b949e}
.blocks-actions{display:flex;gap:8px}

/* ── Paragraph block ── */
.para-block{background:#161b22;border:1px solid #30363d;border-radius:8px;
            margin-bottom:12px;overflow:hidden}
.block-hdr{display:flex;align-items:center;gap:10px;padding:8px 14px;
           background:#1c2128;border-bottom:1px solid #30363d;flex-wrap:wrap}
.block-label{font-size:12px;font-weight:600;color:#8b949e;
             background:#21262d;padding:2px 9px;border-radius:10px;white-space:nowrap}
.block-pills{flex:1;display:flex;flex-wrap:wrap;gap:4px}
.pill{display:inline-block;background:#1c2a3a;color:#79c0ff;padding:2px 9px;
      border-radius:12px;font-size:11px;white-space:nowrap;border:1px solid #1f4068}
.pill .pill-zh{color:#c9d1d9}
.pill .pill-eq{color:#484f58;margin:0 4px}
.no-term{color:#484f58;font-size:11px}
.term-hit{text-decoration:underline dashed #79c0ff;
          text-underline-offset:3px;cursor:help}

/* ── Help / tutorial panel ── */
.help-panel{background:#0d2030;border-bottom:1px solid #1f4068;padding:14px 24px;
            font-size:12.5px;line-height:1.7;color:#8b949e;display:none}
.help-panel.on{display:block}
.help-panel h3{font-size:12px;color:#79c0ff;text-transform:uppercase;
               letter-spacing:.6px;font-weight:600;margin-bottom:8px}
.help-panel ol{margin-left:20px;color:#b0b8c0}
.help-panel ol li{margin-bottom:4px}
.help-panel code{background:#161b22;padding:1px 6px;border-radius:4px;
                 font-family:'SFMono-Regular',Consolas,monospace;font-size:11.5px;
                 color:#79c0ff}
.help-toggle{background:#21262d;color:#8b949e;border:1px solid #30363d;
             border-radius:50%;width:26px;height:26px;cursor:pointer;
             font-size:13px;font-weight:600;line-height:1;padding:0}
.help-toggle:hover{background:#30363d;color:#c9d1d9}
.help-toggle.active{background:#1c2a3a;color:#79c0ff;border-color:#1f4068}

/* ── Merge / unmerge buttons ── */
.btn-merge{padding:4px 10px;font-size:11px;border-radius:5px;border:1px solid #1f4068;
           background:#0d2030;color:#79c0ff;cursor:pointer;font-weight:500;
           white-space:nowrap;transition:background .15s}
.btn-merge:hover:not(:disabled){background:#1c2a3a}
.btn-merge:disabled{opacity:.3;cursor:not-allowed}

/* ── Paragraph divider inside a merged group ── */
.para-divider{padding:5px 14px;color:#79c0ff;font-size:10.5px;
              background:#0d2030;border-top:1px solid #1f4068;
              border-bottom:1px solid #1f4068;letter-spacing:.4px;
              font-weight:500}
.para-block.is-merged{border-color:#1f4068}
.para-block.is-merged .block-hdr{background:#0d2030}

/* ── Sentence rows inside a block ── */
.block-body{width:100%}
.sent-row{display:grid;grid-template-columns:38px 1fr 1fr;border-bottom:1px solid #21262d}
.sent-row:last-child{border-bottom:none}
.sent-row:hover{background:#1c2128}
.sent-num{display:flex;align-items:flex-start;justify-content:center;
          padding:10px 4px;color:#484f58;font-size:11px;font-weight:600;
          border-right:1px solid #21262d;min-width:38px}
.sent-zh{padding:10px 12px;font-size:13px;line-height:1.8;color:#c9d1d9;
         border-right:1px solid #21262d;word-break:break-all}
.sent-en{padding:8px 10px}
.sent-en textarea{width:100%;background:transparent;border:1px solid transparent;
                  border-radius:4px;padding:4px 8px;color:#c9d1d9;font-size:13px;
                  font-family:'SFMono-Regular',Consolas,monospace;resize:none;
                  line-height:1.8;overflow:hidden;min-height:32px;display:block}
.sent-en textarea:hover{border-color:#30363d;background:#0d1117}
.sent-en textarea:focus{border-color:#388bfd;background:#0d1117;outline:none}

/* ── Combined view (per block) ── */
.combined-view{display:none;padding:0}
.combined-body{display:grid;grid-template-columns:1fr 1fr}
.combined-zh{padding:12px 14px;font-size:13px;line-height:1.9;color:#c9d1d9;
             border-right:1px solid #21262d;word-break:break-all}
.sent-sep{display:block;color:#484f58;font-size:11px;margin:4px 0 2px}
.combined-en{padding:10px 12px}
.combined-en textarea{width:100%;background:#0d1117;border:1px solid #30363d;
                      border-radius:6px;padding:10px 12px;color:#c9d1d9;
                      font-size:13px;font-family:'SFMono-Regular',Consolas,monospace;
                      resize:none;line-height:1.9;overflow:hidden;display:block}
.combined-en textarea:focus{border-color:#388bfd;outline:none}

/* ── Preview modal ── */
.overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.82);
         z-index:200;overflow-y:auto;padding:24px 16px}
.overlay.on{display:flex;justify-content:center;align-items:flex-start}
.modal{background:#161b22;border:1px solid #30363d;border-radius:12px;
       width:100%;max-width:980px;margin:auto}
.modal-hdr{padding:14px 20px;border-bottom:1px solid #30363d;
           display:flex;align-items:center;justify-content:space-between}
.modal-hdr h2{font-size:15px;font-weight:600;color:#f0f6fc}
.close-btn{background:none;border:none;color:#8b949e;font-size:22px;
           cursor:pointer;line-height:1;padding:0 4px}
.close-btn:hover{color:#c9d1d9}
.modal-body{padding:20px}
.preview-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.preview-col h3{font-size:11px;color:#8b949e;text-transform:uppercase;
                letter-spacing:.5px;margin-bottom:10px;font-weight:600}
.preview-box{background:#0d1117;border:1px solid #30363d;border-radius:8px;
             padding:14px 16px;font-size:13px;line-height:1.9;
             max-height:62vh;overflow-y:auto;word-break:break-word}
.pv-para{margin-bottom:12px;padding-bottom:12px;border-bottom:1px solid #21262d}
.pv-para:last-child{border-bottom:none;margin-bottom:0}
.modal-ftr{padding:14px 20px;border-top:1px solid #30363d;
           display:flex;gap:10px;justify-content:flex-end}

/* ── Toast ── */
.toast{position:fixed;bottom:22px;right:22px;background:#238636;color:#fff;
       padding:10px 20px;border-radius:8px;font-size:13px;z-index:999;
       display:none;box-shadow:0 4px 14px rgba(0,0,0,.5)}
.toast.on{display:block;animation:slideUp .2s ease}
@keyframes slideUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
</style>
</head>
<body>

<!-- ── Header ── -->
<div class="hdr">
  <h1>🎮 Crowdin Translation Checker</h1>
  <button class="help-toggle" id="helpToggle" onclick="toggleHelp()" title="Show / hide instructions">?</button>
  <div class="gloss-tag loading" id="glossTag">Loading glossary…</div>
</div>

<!-- ── Help / tutorial panel ── -->
<div class="help-panel" id="helpPanel">
  <h3>How to use</h3>
  <ol>
    <li>Paste the <strong>Chinese source</strong> on the left. The <strong>English draft</strong> on the right is <em>optional</em> &mdash; paste one to review/edit, or leave it empty to translate from scratch directly in the table below.</li>
    <li>Use the literal token <code>\n</code> to separate paragraphs &mdash; in both inputs when present, ideally one-to-one. If the English side has fewer (or no) <code>\n</code>, the output will follow the Chinese source's <code>\n</code> structure.</li>
    <li>Click <strong>Build Table</strong>. Each paragraph becomes a row that you can edit sentence by sentence, or toggle <strong>Merge view</strong> to edit a whole paragraph at once.</li>
    <li>Glossary hits are shown as <code>中文 = English</code> pills above each block; matched substrings in the source get a <span class="term-hit">dashed underline</span> (hover for the English term).</li>
    <li><strong>⬇ Merge next</strong> visually combines two adjacent blocks so you can review related paragraphs side by side. <em>This is view-only</em> &mdash; the copied output is unaffected. Press <strong>↑ Unmerge</strong> on a merged block to split it back apart.</li>
    <li><strong>Preview</strong> renders edits live: color codes like <code>[E7594C]…[-]</code> become colored text, and any literal <code>\n</code> you type inside an edit area becomes a real line break. Both are preserved verbatim in the copied output.</li>
    <li>Click <strong>Copy</strong> &mdash; the result always preserves the <em>exact</em> <code>\n</code> structure of your English draft (e.g. <code>\n\n</code> stays <code>\n\n</code>), no matter how you merge / unmerge. Nothing is auto-added.</li>
  </ol>
</div>

<!-- ── Input panel ── -->
<div class="panel">
  <div class="grid2">
    <div class="field">
      <label>Chinese Source <span>(use \n to separate paragraphs; supports [RRGGBB]…[-] color codes)</span></label>
      <textarea id="zh"
        placeholder="Paste Chinese source, e.g.&#10;各位市民即将可以选择离开本城市。但为了维护本城秩序，请遵守规定。\n[E7594C]请注意！[-]活动期间请保持冷静。"></textarea>
    </div>
    <div class="field">
      <label>English Draft <span>(optional — leave empty to translate from scratch; use \n to separate paragraphs, should match the Chinese 1:1)</span></label>
      <textarea id="en"
        placeholder="Optional. Paste an existing English draft to review, or leave empty and fill it in below.&#10;e.g. Citizens will soon be able to leave the city.\n[E7594C]Please note![-] Stay calm during the event."></textarea>
    </div>
  </div>
  <div class="row-btns">
    <button class="btn btn-blue" onclick="buildBlocks()">📊 Build Table</button>
    <div class="status" id="status"></div>
  </div>
</div>

<!-- ── Paragraph blocks ── -->
<div class="blocks-wrap" id="blocksWrap" style="display:none">
  <div class="blocks-hdr">
    <span><strong id="cnt" style="color:#f0f6fc">0</strong> paragraph(s)</span>
    <div class="blocks-actions">
      <button class="btn btn-sky" onclick="showPreview()">👁 Preview</button>
      <button class="btn btn-green" onclick="doCopy()">📋 Copy Final Result</button>
    </div>
  </div>
  <div id="blocks"></div>
</div>

<!-- ── Preview modal ── -->
<div class="overlay" id="overlay">
  <div class="modal">
    <div class="modal-hdr">
      <h2>Preview &mdash; Colors &amp; Paragraph Layout</h2>
      <button class="close-btn" onclick="closePreview()">×</button>
    </div>
    <div class="modal-body">
      <div class="preview-grid">
        <div class="preview-col">
          <h3>Chinese Source</h3>
          <div class="preview-box" id="pvZh"></div>
        </div>
        <div class="preview-col">
          <h3>English (current edit)</h3>
          <div class="preview-box" id="pvEn"></div>
        </div>
      </div>
    </div>
    <div class="modal-ftr">
      <button class="btn btn-gray" onclick="closePreview()">Close</button>
      <button class="btn btn-green" onclick="confirmAndCopy()">✓ Confirm &amp; Copy to Clipboard</button>
    </div>
  </div>
</div>

<!-- ── Toast ── -->
<div class="toast" id="toast">✓ Copied! Paste back into Crowdin.</div>

<script>
// ── State ──────────────────────────────────────────────────────────────────
let glossary = [];

// IMMUTABLE after buildBlocks() — these define the original input structure and
// the final output structure. Merge / unmerge NEVER touches them.
//   paras[i]  = { zhFull, enFull, zhSentences, enSentences }
//   zhSeps[i] / enSeps[i] = literal-\n separator string between paras[i] / paras[i+1]
let paras = [], zhSeps = [], enSeps = [];

// MUTABLE view-layer state. Each group is a visual block on screen.
// Merging adjacent groups concatenates their paraIdxs; unmerging splits back.
//   groups[gi] = { paraIdxs: number[], combined: boolean }
// Initial state: groups.length === paras.length, each holds one paragraph.
let groups = [];

// ── Help toggle ─────────────────────────────────────────────────────────────
function toggleHelp() {
  document.getElementById('helpPanel').classList.toggle('on');
  document.getElementById('helpToggle').classList.toggle('active');
}

// ── Load glossary ───────────────────────────────────────────────────────────
window.addEventListener('load', async () => {
  const tag = document.getElementById('glossTag');
  try {
    const res  = await fetch('/api/glossary');
    const data = await res.json();
    if (data.error) {
      tag.textContent = '⚠ ' + data.error;
      tag.className = 'gloss-tag warn';
    } else {
      glossary = data.terms;
      tag.textContent = '✓ Glossary: ' + glossary.length + ' terms';
      tag.className = 'gloss-tag';
    }
  } catch (e) {
    tag.textContent = '⚠ Failed to load glossary';
    tag.className = 'gloss-tag warn';
  }
});

// ── Helpers ─────────────────────────────────────────────────────────────────
function colorize(text) {
  return text.replace(/\[([0-9A-Fa-f]{6})\]([\s\S]*?)\[-\]/g,
    (_, hex, body) => '<span style="color:#' + hex + '">' + body + '</span>');
}

function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Convert literal "\n" (backslash + n, two chars) typed by the user inside an
// edit textarea into a real <br> for rendered HTML output. Safe to apply AFTER
// esc()/highlightTerms() because neither emits nor escapes a backslash, so the
// only "\n" sequences left in the html are the user's own.
function nlToBr(html) {
  return html.replace(/\\n/g, '<br>');
}

function autoH(ta) {
  ta.style.height = 'auto';
  ta.style.height = (ta.scrollHeight + 2) + 'px';
}

function setStatus(msg, type) {
  const el = document.getElementById('status');
  el.textContent = msg;
  el.className = 'status ' + type;
}

function findTerms(zhText) {
  return glossary.filter(({zh}) => zhText.includes(zh));
}

// Wrap matched glossary terms in <span class="term-hit"> while escaping HTML.
// Operates on raw text so we can safely interleave escaped chunks with markup.
// Longer terms win when overlapping (e.g. "首领部队" beats "首领").
function highlightTerms(rawText, terms) {
  if (!terms || !terms.length) return esc(rawText);

  const sorted = [...terms].sort((a, b) => b.zh.length - a.zh.length);
  const matches = [];
  for (const t of sorted) {
    let idx = 0;
    while ((idx = rawText.indexOf(t.zh, idx)) !== -1) {
      const start = idx, end = idx + t.zh.length;
      const overlaps = matches.some(m => start < m.end && end > m.start);
      if (!overlaps) matches.push({ start, end, term: t });
      idx = end;
    }
  }
  matches.sort((a, b) => a.start - b.start);

  let html = '', cursor = 0;
  for (const m of matches) {
    html += esc(rawText.slice(cursor, m.start));
    html += '<span class="term-hit" title="' + esc(m.term.en) + '">'
         +  esc(rawText.slice(m.start, m.end))
         +  '</span>';
    cursor = m.end;
  }
  html += esc(rawText.slice(cursor));
  return html;
}

// Render ZH text with term highlighting + literal \n -> <br> + color codes.
// Pipeline order matters: highlightTerms wraps escaped chunks, nlToBr converts
// any user-typed \n inside the body, and colorize finally wraps [RRGGBB]…[-].
function renderZh(text) {
  return colorize(nlToBr(highlightTerms(text, findTerms(text))));
}

// Mask [RRGGBB]...[‐] spans so their internal punctuation doesn't trigger splits.
// Returns masked string (same length, placeholder chars).
function maskColors(text) {
  return text.replace(/\[[0-9A-Fa-f]{6}\][\s\S]*?\[-\]/g,
    m => '\x01'.repeat(m.length));
}

// Split ZH paragraph into sentences, keeping [color]...[‐] spans atomic.
// Terminators: 。！？
function splitZhSentences(text) {
  if (!text.trim()) return [text];
  const masked = maskColors(text);
  const parts = [];
  // \x01 is NOT a terminator; runs of it stay with surrounding text
  const re = /[^。！？]+[。！？]*/g;
  let m;
  while ((m = re.exec(masked)) !== null) {
    const slice = text.slice(m.index, m.index + m[0].length).trim();
    if (slice) parts.push(slice);
  }
  return parts.length ? parts : [text.trim()];
}

// Split EN paragraph into sentences.
// Only split at .!? that is followed by whitespace + an uppercase letter.
// This avoids false splits on decimals (0.1%), abbreviations, color codes, etc.
//
// Two-pass strategy:
//   1) Split on the standard "<.!?> + whitespace + UPPERCASE" boundary.
//   2) Re-glue tiny list-marker fragments (e.g. "1.", "2.", "i.", "I.") onto
//      the next sentence so "1. During the event..." stays as one sentence
//      instead of becoming ["1.", "During the event..."].
function splitEnSentences(text) {
  if (!text.trim()) return [text];
  const raw = text.split(/(?<=[.!?])\s+(?=[A-Z])/).map(s => s.trim()).filter(Boolean);
  const isListMarker = s => /^[0-9]{1,3}\.$/.test(s) || /^[ivxIVX]{1,4}\.$/.test(s);

  const out = [];
  let pending = '';
  for (const p of raw) {
    if (isListMarker(p)) {
      pending = pending ? pending + ' ' + p : p;
    } else {
      out.push(pending ? pending + ' ' + p : p);
      pending = '';
    }
  }
  if (pending) out.push(pending);
  return out;
}

// Parse raw Crowdin input into non-empty paragraphs PLUS the exact \n separator
// between each adjacent pair. Splitting on literal \n yields tokens; runs of empty
// tokens in between encode multi-\n separators (e.g. "A\n\nB" -> ["A","","B"] = sep "\n\n").
// Returns { paras: string[], seps: string[] } where seps.length === paras.length - 1.
function parseRawWithSeps(raw) {
  const tokens = raw.split('\\n');
  const out = [], seps = [];
  let pendingEmpty = 0;
  for (const tok of tokens) {
    const trimmed = tok.trim();
    if (trimmed) {
      if (out.length > 0) {
        // Separator = (1 boundary \n) + (one extra \n per empty token in between)
        seps.push('\\n'.repeat(1 + pendingEmpty));
      }
      out.push(trimmed);
      pendingEmpty = 0;
    } else {
      pendingEmpty++;
    }
  }
  return { paras: out, seps };
}

// ── Build blocks ─────────────────────────────────────────────────────────────
function buildBlocks() {
  const zhRaw = document.getElementById('zh').value.trim();
  const enRaw = document.getElementById('en').value.trim();
  if (!zhRaw) { alert('Please paste the Chinese source.'); return; }
  // English draft is OPTIONAL — leave it empty to translate from scratch using
  // the table's per-sentence textareas.

  const zh = parseRawWithSeps(zhRaw);
  const en = parseRawWithSeps(enRaw);
  const n  = Math.max(zh.paras.length, en.paras.length);

  paras = [];
  for (let i = 0; i < n; i++) {
    const zhFull = zh.paras[i] || '';
    const enFull = en.paras[i] || '';
    paras.push({
      zhFull,
      enFull,
      zhSentences: splitZhSentences(zhFull),
      enSentences: splitEnSentences(enFull)
    });
  }

  // Fill separators for the n-1 gaps. EN can be partial / empty; when no EN
  // separator is available, fall back to the ZH separator so the eventual
  // copied output mirrors the source's \n structure rather than collapsing it.
  zhSeps = [];
  enSeps = [];
  for (let i = 0; i < n - 1; i++) {
    const zSep = zh.seps[i] !== undefined ? zh.seps[i] : '\\n';
    const eSep = en.seps[i] !== undefined ? en.seps[i] : zSep;
    zhSeps.push(zSep);
    enSeps.push(eSep);
  }

  // Reset view-layer: every paragraph starts as its own group
  groups = paras.map((_, i) => ({ paraIdxs: [i], combined: false }));

  document.getElementById('cnt').textContent = n;

  // Show wrapper BEFORE renderBlocks so textareas have non-zero scrollHeight
  // when autoH measures them (otherwise heights collapse to min-height = 1 line
  // and long EN content gets clipped by overflow:hidden).
  const wrap = document.getElementById('blocksWrap');
  wrap.style.display = 'block';
  renderBlocks();
  wrap.scrollIntoView({ behavior: 'smooth', block: 'start' });
  setStatus('Built ' + n + ' paragraph(s) ✓', 'ok');
}

// VIEW-LAYER ONLY. Merge group `gi` with group `gi+1` so they render as one
// visual block. Original paras/seps are untouched -> Copy & Preview output the
// EXACT \n structure of the user's draft regardless of how many merges happen.
function mergeWithNext(gi) {
  if (gi < 0 || gi >= groups.length - 1) return;
  groups[gi].paraIdxs = [...groups[gi].paraIdxs, ...groups[gi + 1].paraIdxs];
  groups[gi].combined = false; // combined-view doesn't make sense for merged groups
  groups.splice(gi + 1, 1);
  renderBlocks();
  const labels = groups[gi].paraIdxs.map(i => i + 1).join(' + ');
  setStatus('View-merged paragraphs ' + labels + ' ✓ (output unchanged)', 'ok');
}

// VIEW-LAYER ONLY. Split a multi-paragraph group back into individual groups.
function unmergeGroup(gi) {
  if (gi < 0 || gi >= groups.length) return;
  const g = groups[gi];
  if (g.paraIdxs.length < 2) return;
  const expanded = g.paraIdxs.map(idx => ({ paraIdxs: [idx], combined: false }));
  groups.splice(gi, 1, ...expanded);
  renderBlocks();
  setStatus('Unmerged ✓', 'ok');
}

function renderBlocks() {
  const container = document.getElementById('blocks');
  container.innerHTML = '';
  groups.forEach((g, gi) => container.appendChild(makeBlock(g, gi)));
  // rAF ensures layout has settled before reading scrollHeight
  requestAnimationFrame(() => {
    container.querySelectorAll('textarea').forEach(autoH);
  });
}

// Render one visual block per group. A group with N paragraphs renders all N
// paragraphs' sentences sequentially with a divider between them.
function makeBlock(g, gi) {
  const isMulti = g.paraIdxs.length > 1;
  const isLast  = (gi === groups.length - 1);

  // ── Aggregate glossary hits across all paragraphs in this group (deduped)
  const seen  = new Set();
  const terms = [];
  for (const pIdx of g.paraIdxs) {
    for (const t of findTerms(paras[pIdx].zhFull)) {
      const key = t.zh + '\x01' + t.en;
      if (!seen.has(key)) { seen.add(key); terms.push(t); }
    }
  }
  const termHtml = terms.length
    ? terms.map(t =>
        '<span class="pill">'
      +   '<span class="pill-zh">' + esc(t.zh) + '</span>'
      +   '<span class="pill-eq">=</span>'
      +   esc(t.en)
      + '</span>'
      ).join('')
    : '<span class="no-term">no glossary hits</span>';

  // ── Sentence rows. Walk every paragraph in the group; insert a divider
  //    before each paragraph after the first so users can still see the
  //    original boundary. Each textarea is bound to its ORIGINAL paragraph
  //    index via data-para so edits are attributed correctly even after merge.
  let sentRows = '';
  let rowNum   = 0;
  g.paraIdxs.forEach((pIdx, idxInGroup) => {
    const p = paras[pIdx];
    if (idxInGroup > 0) {
      sentRows += `<div class="para-divider">— Paragraph ${pIdx + 1} —</div>`;
    }
    const rowCount = Math.max(p.zhSentences.length, p.enSentences.length, 1);
    for (let si = 0; si < rowCount; si++) {
      rowNum++;
      const zhVal = p.zhSentences[si] !== undefined ? p.zhSentences[si] : '';
      const enVal = p.enSentences[si] !== undefined ? p.enSentences[si] : '';
      sentRows += `<div class="sent-row">
        <div class="sent-num">${rowNum}</div>
        <div class="sent-zh">${zhVal ? renderZh(zhVal) : ''}</div>
        <div class="sent-en"><textarea
          data-para="${pIdx}" data-sent="${si}"
          oninput="autoH(this);syncCombined(${gi})">${esc(enVal)}</textarea></div>
      </div>`;
    }
  });

  // ── Combined view (works for both single and merged groups). Sentence-level
  //    textareas always exist in the DOM (just hidden behind combined view) and
  //    keep their data-para attribution, so syncSentences distributes re-split
  //    sentences back into the correct ORIGINAL paragraphs by DOM order.
  const zhSentencesAll = g.paraIdxs.flatMap(pIdx => paras[pIdx].zhSentences);
  const zhCombined = zhSentencesAll.length > 1
    ? zhSentencesAll.map((s, si) =>
        (si > 0 ? '<span class="sent-sep"> </span>' : '') + renderZh(s)
      ).join('')
    : renderZh(zhSentencesAll[0] || '');
  const enJoined = g.paraIdxs.map(pIdx => paras[pIdx].enFull).filter(Boolean).join(' ');
  const combinedHtml = `
    <div class="combined-view" id="body-comb-${gi}">
      <div class="combined-body">
        <div class="combined-zh">${zhCombined}</div>
        <div class="combined-en">
          <textarea data-combined="1"
            oninput="autoH(this);syncSentences(${gi})">${esc(enJoined)}</textarea>
        </div>
      </div>
    </div>`;

  // ── Header: label + pills + buttons
  const labelText = isMulti
    ? 'Paragraphs ' + g.paraIdxs.map(i => i + 1).join(' + ')
    : 'Paragraph ' + (g.paraIdxs[0] + 1);
  const unmergeBtn = isMulti
    ? `<button class="btn-merge" onclick="unmergeGroup(${gi})"
         title="Split this merged group back into individual paragraphs">↑ Unmerge</button>`
    : '';
  const mergeBtn = `<button class="btn-merge" onclick="mergeWithNext(${gi})" ${isLast ? 'disabled' : ''}
       title="Merge with next block (view-only — does not change the copied output)">⬇ Merge next</button>`;
  const toggleBtn = `<button class="btn-sm" id="toggle-${gi}" onclick="toggleBlock(${gi})">Merge view ▾</button>`;

  const div = document.createElement('div');
  div.className = 'para-block' + (isMulti ? ' is-merged' : '');
  div.id = 'block-' + gi;
  div.innerHTML = `
    <div class="block-hdr">
      <span class="block-label">${labelText}</span>
      <div class="block-pills">${termHtml}</div>
      ${unmergeBtn}
      ${mergeBtn}
      ${toggleBtn}
    </div>
    <div class="block-body" id="body-sent-${gi}">${sentRows}</div>
    ${combinedHtml}`;

  return div;
}

// ── Toggle a group between sentence-level rows and a single combined textarea.
//    Works for both single-paragraph and merged groups; in merged groups the
//    combined textarea spans all paragraphs and on edit-commit the resulting
//    sentences flow back into the correct ORIGINAL paragraphs by DOM order
//    (each sent-row textarea preserves its data-para attribution).
function toggleBlock(gi) {
  const g = groups[gi];
  if (!g) return;
  g.combined = !g.combined;

  const sentView = document.getElementById('body-sent-' + gi);
  const combView = document.getElementById('body-comb-' + gi);
  const btn      = document.getElementById('toggle-' + gi);

  if (g.combined) {
    const sentTas = sentView.querySelectorAll('textarea');
    const joined  = Array.from(sentTas).map(ta => ta.value).filter(Boolean).join(' ');
    const combTa  = combView.querySelector('textarea');
    combTa.value  = joined;
    sentView.style.display = 'none';
    combView.style.display = 'block';
    btn.textContent = 'Split view ▴';
    btn.classList.add('active');
    autoH(combTa);
  } else {
    const combTa  = combView.querySelector('textarea');
    const reSplit = splitEnSentences(combTa.value);
    const sentTas = sentView.querySelectorAll('textarea');
    sentTas.forEach((ta, si) => {
      ta.value = reSplit[si] !== undefined ? reSplit[si] : '';
    });
    sentView.style.display = 'block';
    combView.style.display = 'none';
    btn.textContent = 'Merge view ▾';
    btn.classList.remove('active');
    sentTas.forEach(autoH);
  }
}

// Sync combined textarea when a sentence textarea changes (single-para groups).
function syncCombined(gi) {
  const combTa = document.querySelector('#body-comb-' + gi + ' textarea');
  if (!combTa) return; // multi-para groups have no combined view
  const sentTas = document.querySelectorAll('#body-sent-' + gi + ' textarea');
  combTa.value = Array.from(sentTas).map(ta => ta.value).filter(Boolean).join(' ');
  autoH(combTa);
}

// Re-split combined EN into sentence rows when combined textarea changes.
function syncSentences(gi) {
  const combTa  = document.querySelector('#body-comb-' + gi + ' textarea');
  if (!combTa) return;
  const reSplit = splitEnSentences(combTa.value);
  const sentTas = document.querySelectorAll('#body-sent-' + gi + ' textarea');
  sentTas.forEach((ta, si) => {
    ta.value = reSplit[si] !== undefined ? reSplit[si] : '';
    autoH(ta);
  });
}

// ── Collect final EN for ORIGINAL paragraph index `pIdx`, regardless of which
//    group currently displays it. Sentence-level textareas are always present
//    in the DOM (even when hidden behind combined view) and are kept in sync
//    by syncSentences, so they are the single source of truth.
function getParaEN(pIdx) {
  const tas = document.querySelectorAll(
    'textarea[data-para="' + pIdx + '"][data-sent]'
  );
  return Array.from(tas).map(ta => ta.value).filter(Boolean).join(' ');
}

// ── Preview modal ────────────────────────────────────────────────────────────
// Always renders the ORIGINAL paragraph structure with the EXACT \n separators
// from the user's draft. Merge / unmerge are pure view-layer operations and
// have NO effect here.
function showPreview() {
  if (!paras.length) { alert('Please build the table first.'); return; }

  const nlCount = sep => (sep.match(/\\n/g) || []).length;

  let zhHtml = '';
  let enHtml = '';
  paras.forEach((p, i) => {
    zhHtml += renderZh(p.zhFull);
    // EN side may contain user-typed color codes [RRGGBB]…[-] AND literal \n.
    // Pipeline: esc -> nlToBr -> colorize so all three get rendered live.
    enHtml += colorize(nlToBr(esc(getParaEN(i))));
    if (i < paras.length - 1) {
      zhHtml += '<br>'.repeat(nlCount(zhSeps[i]));
      enHtml += '<br>'.repeat(nlCount(enSeps[i]));
    }
  });
  document.getElementById('pvZh').innerHTML = zhHtml;
  document.getElementById('pvEn').innerHTML = enHtml;

  document.getElementById('overlay').classList.add('on');
}

function closePreview() {
  document.getElementById('overlay').classList.remove('on');
}

document.getElementById('overlay').addEventListener('click', e => {
  if (e.target === document.getElementById('overlay')) closePreview();
});

// ── Copy: rebuild the result using the ORIGINAL paragraph list and the EXACT
//         \n separators captured from the user's draft. Merge / unmerge are
//         purely visual and never touch this output. Nothing is auto-inserted.
async function doCopy() {
  let result = '';
  paras.forEach((_, i) => {
    result += getParaEN(i);
    if (i < paras.length - 1) result += enSeps[i];
  });

  try {
    await navigator.clipboard.writeText(result);
  } catch {
    const ta = document.createElement('textarea');
    ta.value = result;
    ta.style.cssText = 'position:fixed;opacity:0;top:0;left:0';
    document.body.appendChild(ta);
    ta.focus(); ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
  }
  const toast = document.getElementById('toast');
  toast.classList.add('on');
  setTimeout(() => toast.classList.remove('on'), 3000);
}

function confirmAndCopy() { doCopy(); closePreview(); }
</script>
</body>
</html>"""


# ─────────────────────────────────────────────────────────────────────────────
# HTTP request handler
# ─────────────────────────────────────────────────────────────────────────────
class Handler(http.server.BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        pass

    def _send_json(self, code: int, data: dict):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/":
            body = HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif self.path == "/api/glossary":
            terms, msg = load_glossary()
            if terms:
                self._send_json(200, {"terms": terms, "message": msg})
            else:
                self._send_json(200, {"terms": [], "error": msg})
        else:
            self.send_error(404)


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────
class _Server(http.server.HTTPServer):
    allow_reuse_address = True


if __name__ == "__main__":
    # Kill any process still holding the port
    subprocess.run(f"lsof -ti:{PORT} | xargs kill -9 2>/dev/null; true",
                   shell=True, check=False)
    time.sleep(0.3)

    server = _Server(("127.0.0.1", PORT), Handler)
    print(f"Crowdin Checker  →  http://localhost:{PORT}")
    print(f"Glossary: {os.path.join(SCRIPT_DIR, GLOSSARY_FILE)}")
    print("Press Ctrl+C to stop.\n")

    def _open_browser():
        time.sleep(0.7)
        webbrowser.open(f"http://localhost:{PORT}")

    threading.Thread(target=_open_browser, daemon=True).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
