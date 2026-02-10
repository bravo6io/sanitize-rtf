#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

// Usage:
//   node batch_simplify_rtf.js --dir /path/to/rtfs --styles styles.tsv
//
// For each .rtf file in the directory (excluding *_sanitized.rtf), this script
// runs simplify_rtf.js and writes "_sanitized" outputs, overwriting if present.

function usage() {
  console.error('Usage: node batch_simplify_rtf.js --dir DIR [--styles styles.tsv] [--caps]');
  process.exit(1);
}

function parseArgs(argv) {
  const args = { dir: '', styles: 'styles.tsv', caps: false };
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--dir') args.dir = argv[++i] || '';
    else if (a === '--styles') args.styles = argv[++i] || '';
    else if (a.toLowerCase() === '--caps') args.caps = true;
    else usage();
  }
  if (!args.dir) usage();
  return args;
}

function listRtfFiles(dir) {
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  return entries
    .filter((e) => e.isFile())
    .map((e) => e.name)
    .filter((name) => name.toLowerCase().endsWith('.rtf'))
    .filter((name) => !name.toLowerCase().endsWith('_sanitized.rtf'))
    .map((name) => path.join(dir, name));
}

function outputPathFor(inputPath) {
  const dir = path.dirname(inputPath);
  const ext = path.extname(inputPath);
  const base = path.basename(inputPath, ext);
  return path.join(dir, `${base}_sanitized${ext}`);
}

function runSimplify(inputPath, outputPath, stylesPath, caps) {
  const script = path.join(__dirname, 'simplify_rtf.js');
  const args = [script, '--in', inputPath, '--out', outputPath, '--styles', stylesPath];
  if (caps) args.push('--caps');
  const result = spawnSync('node', args, {
    stdio: 'inherit',
  });
  if (result.status !== 0) {
    throw new Error(`Failed on ${path.basename(inputPath)} (exit ${result.status})`);
  }
}

function main() {
  const args = parseArgs(process.argv);
  const dir = path.resolve(args.dir);
  const stylesPath = path.resolve(args.styles);

  if (!fs.existsSync(dir) || !fs.statSync(dir).isDirectory()) {
    throw new Error(`Not a directory: ${dir}`);
  }
  if (!fs.existsSync(stylesPath)) {
    throw new Error(`Styles file not found: ${stylesPath}`);
  }

  const files = listRtfFiles(dir);
  if (files.length === 0) {
    console.log('No .rtf files to process.');
    return;
  }

  for (const inputPath of files) {
    const outPath = outputPathFor(inputPath);
    runSimplify(inputPath, outPath, stylesPath, args.caps);
  }
}

try {
  main();
} catch (err) {
  console.error(err.message || String(err));
  process.exit(1);
}
