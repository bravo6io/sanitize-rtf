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
  s = s.replace(/[^\x20-\x7E]/g, '');
  s = s.replace(/[ ]{2,}/g, ' ');
  s = s.replace(/\s+([,.;:])/g, '$1');
  return s.trim();
}

// Remove any starred destination groups (e.g. {\*\themedata ...}).
function stripStarGroups(rtf) {
  let out = '';
  let i = 0;
  while (i < rtf.length) {
    if (rtf[i] === '{' && rtf[i + 1] === '\\' && rtf[i + 2] === '*') {
      let depth = 1;
      i += 3;
      while (i < rtf.length && depth > 0) {
        if (rtf[i] === '{') depth++;
        else if (rtf[i] === '}') depth--;
        i++;
      }
      continue;
    }
    out += rtf[i++];
  }
  return out;
}

// Remove non-content groups (font tables, styles, info, etc) from RTF.
function stripGroupsByKeyword(rtf, keywords) {
  let out = '';
  let i = 0;
  while (i < rtf.length) {
    if (rtf[i] === '{') {
      let j = i + 1;
      if (rtf[j] === '\\') {
        j++;
        if (rtf[j] === '*') j++;
        const start = j;
        while (j < rtf.length && /[a-zA-Z]/.test(rtf[j])) j++;
        const word = rtf.slice(start, j);
        if (keywords.has(word)) {
          let depth = 1;
          i = j;
          while (i < rtf.length && depth > 0) {
            if (rtf[i] === '{') depth++;
            else if (rtf[i] === '}') depth--;
            i++;
          }
          continue;
        }
      }
    }
    out += rtf[i++];
  }
  return out;
}

// Split RTF into paragraph fragments using \par control words only.
function splitParagraphs(rtf) {
  const paras = [];
  const re = /\\par(?![a-zA-Z])/g;
  let last = 0;
  let m;
  while ((m = re.exec(rtf)) !== null) {
    paras.push(rtf.slice(last, m.index));
    last = re.lastIndex;
  }
  if (last < rtf.length) paras.push(rtf.slice(last));
  return paras;
}

// Extract ordered entries (list items and normal text) from RTF.
function extractEntriesFromRtf(rtf) {
  const entries = [];
  const stripKeywords = new Set([
    'fonttbl',
    'stylesheet',
    'info',
    'colortbl',
    'listtable',
    'listoverridetable',
    'rsidtbl',
    'xmlnstbl',
    'datastore',
    'themedata',
    'colorschememapping',
    'latentstyles',
    'generator',
  ]);
  const noStarGroups = stripStarGroups(rtf.replace(/\r\n/g, '\n'));
  const cleaned = stripGroupsByKeyword(noStarGroups, stripKeywords);
  const paras = splitParagraphs(cleaned);
  const orderedLabelRe = /([0-9]+\.|[a-z]\.)\\tab/i;

  function extractListtextGroup(s) {
    let i = s.indexOf('{\\listtext');
    if (i === -1) return '';
    let depth = 0;
    let start = i;
    for (; i < s.length; i++) {
      if (s[i] === '{') depth++;
      else if (s[i] === '}') {
        depth--;
        if (depth === 0) return s.slice(start, i + 1);
      }
    }
    return '';
  }

  function removeListtextGroups(s) {
    let out = '';
    let i = 0;
    while (i < s.length) {
      if (s.startsWith('{\\listtext', i)) {
        let depth = 1;
        i += 1;
        while (i < s.length && depth > 0) {
          if (s[i] === '{') depth++;
          else if (s[i] === '}') depth--;
          i++;
        }
        continue;
      }
      out += s[i++];
    }
    return out;
  }

  function getListLabel(listtextGroup) {
    if (!listtextGroup) return null;
    if (/\\'b7\\tab/i.test(listtextGroup)) return { label: "\\u8226?", type: 'bullet' };
    const uMatch = listtextGroup.match(/\\u(-?\d+)\??\\tab/i);
    if (uMatch) {
      const code = parseInt(uMatch[1], 10);
      if (code === 8226 || code === 9679) return { label: "\\u8226?", type: 'bullet' };
    }
    const oMatch = listtextGroup.match(orderedLabelRe);
    if (oMatch) return { label: oMatch[1], type: 'ordered' };
    return null;
  }

  for (const para of paras) {
    if (!para || !para.trim()) continue;
    const listIdx = para.search(/\\listtext\b/);
    const insIdx = para.search(/\\insrsid\d+/);
    let contentPara;
    if (listIdx !== -1) {
      let start = listIdx;
      if (start > 0 && para[start - 1] === '{') start -= 1;
      contentPara = para.slice(start);
    } else if (insIdx !== -1) {
      contentPara = para.slice(insIdx);
    } else {
      contentPara = para;
    }
    const listtextGroup = extractListtextGroup(contentPara);
    const labelInfo = getListLabel(listtextGroup);
    if (labelInfo) {
      const label = labelInfo.label;
      const bodyRtf = removeListtextGroups(contentPara);
      const body = sanitizeRtfFragment(bodyRtf);
      if (body) entries.push({ type: 'list', label, body, listType: labelInfo.type });
      continue;
    }
    const text = sanitizeRtfFragment(contentPara);
    if (text) entries.push({ type: 'text', body: text });
  }

  return entries;
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
  const entries = extractEntriesFromRtf(rtf);

  const out = [];
  for (const hl of headerLines) out.push(hl);
  for (const entry of entries) {
    if (entry.type === 'list') {
      const isBullet = entry.listType === 'bullet';
      const isLevel2 = isBullet || /^[a-z]\./i.test(entry.label);
      const prefix = isLevel2 ? style.level2_prefix : style.level1_prefix;
      const labelSep = isLevel2 ? style.level2_label_sep : style.level1_label_sep;
      const label = isBullet && style.bullet_label ? style.bullet_label : entry.label;
      const body = escapeRtfText(entry.body);
      if (!body) continue;
      out.push(prefix + label + labelSep + body + '\\par');
      continue;
    }
    const normalPrefix = style.normal_prefix || '\\pard\\s0 ';
    const body = escapeRtfText(entry.body);
    if (!body) continue;
    out.push(normalPrefix + body + '\\par');
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
