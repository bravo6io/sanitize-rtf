# RTF Sanitizer

This project extracts list items from Word-generated RTF files, strips styling from the text, and re-applies a clean, consistent style defined in `styles.tsv`.

## Purpose
This repo takes `.rtf` output from text editors and simplifies it so SketchUp Layout can render lists reliably for construction document notes. SketchUp Layout’s RTF support is fragile with complex styling, and Word often injects extra formatting that causes misinterpretation. Common issues this repo avoids:
- misaligned lists and sub-lists
- random restarting of numbering
- inconsistent bolding of numbering vs. text in the first line of a paragraph

## Files
- `simplify_rtf.js`: main sanitizer/formatter (Node.js)
- `batch_simplify_rtf.js`: batch processor for directories
- `styles.tsv`: style definitions exported from `simplified.rtf`
- `simplified.rtf`: style reference and target formatting

## Single File Usage
```bash
node simplify_rtf.js --in word.rtf --out output.rtf --styles styles.tsv
```

Notes:
- Input file is unchanged.
- Output file is overwritten if it already exists.

## Batch Usage (Directory)
```bash
node batch_simplify_rtf.js --dir /path/to/rtfs --styles styles.tsv
```

Behavior:
- Processes every `.rtf` file in the directory.
- Skips files ending in `_sanitized.rtf`.
- Writes `<original>_sanitized.rtf` next to each input.
- Overwrites existing `_sanitized.rtf` files.

## Global Command (bash)
A wrapper script `sanirtf` was installed at `~/bin/sanirtf` and `~/bin` was added to `PATH` in `~/.bashrc`.

```bash
sanirtf -dir /path/to/rtfs
```

Optional:
```bash
sanirtf -dir /path/to/rtfs -styles /path/to/styles.tsv
```

If `sanirtf` is not found in your current terminal, run:
```bash
source ~/.bashrc
```

## Fresh Clone Setup (Global PATH)
If someone pulls this repo fresh, they can create the `sanirtf` command like this:

```bash
mkdir -p "$HOME/bin"
cat > "$HOME/bin/sanirtf" <<'SH'
#!/usr/bin/env bash
set -euo pipefail

SCRIPT="/path/to/your/clone/batch_simplify_rtf.js"

# Map short flags to long flags for convenience.
args=()
for a in "$@"; do
  if [[ "$a" == "-dir" ]]; then
    args+=("--dir")
  elif [[ "$a" == "-styles" ]]; then
    args+=("--styles")
  else
    args+=("$a")
  fi
done

exec node "$SCRIPT" "${args[@]}"
SH
chmod +x "$HOME/bin/sanirtf"

# Ensure ~/bin is on PATH
if ! grep -q 'export PATH="$HOME/bin:$PATH"' "$HOME/.bashrc" 2>/dev/null; then
  echo 'export PATH="$HOME/bin:$PATH"' >> "$HOME/.bashrc"
fi
source ~/.bashrc
```

Replace `/path/to/your/clone` with the actual path to this repo.

## Style Source
`styles.tsv` was generated from `simplified.rtf` and contains:
- Header lines (RTF preamble)
- Prefixes for level 1 and level 2 list items
- Label separator (`\tab`)

If you update `simplified.rtf`, regenerate `styles.tsv` before running the sanitizer.
