#!/usr/bin/env node
'use strict';

const fs = require('fs');

// Usage (direct with Node.js):
//   node simplify_rtf.js --in word.rtf --out output.rtf --styles styles.tsv
//
// This script:
// - Reads styles from styles.tsv (exported from simplified.rtf)
// - Extracts numbered/lettered list items from Word RTF
// - Strips all styling from item bodies
// - Emits a clean RTF using the styles.tsv prefixes

// Print usage and exit with failure.
function usage() {
  console.error('Usage: node simplify_rtf.js --in INPUT.rtf --out OUTPUT.rtf [--styles styles.tsv]');
  process.exit(1);
}

// Parse CLI arguments into an options object.
function parseArgs(argv) {
  const args = { in: '', out: '', styles: 'styles.tsv' };
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--in') args.in = argv[++i] || '';
    else if (a === '--out') args.out = argv[++i] || '';
    else if (a === '--styles') args.styles = argv[++i] || '';
    else usage();
  }
  if (!args.in || !args.out) usage();
  return args;
}

// Load header lines and style prefixes from styles.tsv.
function readStyles(path) {
  const lines = fs.readFileSync(path, 'utf8').split(/\n/);
  const headerLines = [];
  const style = {};
  for (const raw of lines) {
    const line = raw.replace(/\r$/, '');
    if (!line.trim()) continue;
    const [type, name, value] = line.split(/\t/, 3);
    if (type === 'type' && name === 'name') continue;
    if (type === 'header_line') headerLines.push(value ?? '');
    if (type === 'style') style[name] = value ?? '';
  }
  for (const req of ['level1_prefix', 'level1_label_sep', 'level2_prefix', 'level2_label_sep']) {
    if (!(req in style)) throw new Error(`Missing style: ${req} in ${path}`);
  }
  if (headerLines.length === 0) throw new Error(`No header_line entries in ${path}`);
  return { headerLines, style };
}

// Convert RTF \uN escapes to ASCII (best-effort).
function decodeUnicode(n) {
  let code = n;
  if (code < 0) code += 65536;
  if (code === 8220 || code === 8221) return '"';
  if (code === 8216 || code === 8217) return "'";
  if (code === 8211 || code === 8212) return '-';
  if (code >= 0 && code < 128) return String.fromCharCode(code);
  return '';
}

// Reduce an RTF fragment to plain ASCII text, removing all RTF control codes.
function sanitizeRtfFragment(fragment) {
  // Reduce an RTF fragment to plain ASCII text, removing all RTF control codes.
  let s = fragment;
  s = s.replace(/\\u(-?\d+)\??/g, (_, n) => decodeUnicode(parseInt(n, 10)));
  s = s.replace(/\\'([0-9a-fA-F]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
  s = s.replace(/\\[a-zA-Z]+-?\d* ?/g, '');
  s = s.replace(/\\[^a-zA-Z]/g, '');
  s = s.replace(/[{}]/g, '');
  s = s.replace(/[\n\r\t]+/g, '');
  s = s.replace(/[ ]{2,}/g, ' ');
  s = s.replace(/\s+([,.;:])/g, '$1');
  return s.trim();
}

// Find list labels like "1.\tab" or "a.\tab" and capture text up to \par.
function extractItemsFromRtf(rtf) {
  // Find list labels like "1.\tab" or "a.\tab" and capture text up to \par.
  const items = [];
  const s = rtf.replace(/\r\n/g, '\n');
  const re = /([0-9]+\.|[a-z]\.)\\tab ?/gi;
  let m;
  while ((m = re.exec(s)) !== null) {
    const label = m[1];
    const start = re.lastIndex;
    const end = s.indexOf('\\par', start);
    if (end === -1) break;
    const bodyRtf = s.slice(start, end);
    const body = sanitizeRtfFragment(bodyRtf);
    if (body) items.push({ label, body });
  }
  return items;
}

// Escape text for safe inclusion in RTF output.
function escapeRtfText(s) {
  let out = s.replace(/([\\{}])/g, '\\$1');
  out = out.replace(/[^\x00-\x7F]/g, '');
  out = out.replace(/[ ]{2,}/g, ' ').trim();
  return out;
}

// Orchestrate reading inputs, extracting items, and writing output.
function main() {
  const args = parseArgs(process.argv);
  const { headerLines, style } = readStyles(args.styles);
  const rtf = fs.readFileSync(args.in, 'utf8');
  const items = extractItemsFromRtf(rtf);

  const out = [];
  for (const hl of headerLines) out.push(hl);
  for (const it of items) {
    const isLevel2 = /^[a-z]\./i.test(it.label);
    const prefix = isLevel2 ? style.level2_prefix : style.level1_prefix;
    const labelSep = isLevel2 ? style.level2_label_sep : style.level1_label_sep;
    const body = escapeRtfText(it.body);
    if (!body) continue;
    out.push(prefix + it.label + labelSep + body + '\\par');
  }
  out.push('}');
  fs.writeFileSync(args.out, out.join('\n'));
}

try {
  main();
} catch (err) {
  console.error(err.message || String(err));
  process.exit(1);
}
