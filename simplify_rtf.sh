#!/usr/bin/env bash
#
# simplify_rtf.sh
# Usage:
#   ./simplify_rtf.sh --in INPUT.rtf --out OUTPUT.rtf
#
# What it does:
# - Reads a Word-generated RTF file
# - Extracts list items from \listtext entries
# - Writes a minimal RTF using Roboto 12pt with:
#   - Level 1: "1. 2. 3. ..."
#   - Level 2: "a. b. c. ..." (indented with \tab)
#
# Note:
# - Requires gawk for strtonum().

set -euo pipefail

if [[ $# -lt 4 ]]; then
  echo "Usage: $0 --in INPUT.rtf --out OUTPUT.rtf" >&2
  exit 1
fi

in=""
out=""
while [[ $# -gt 0 ]]; do
  case "$1" in
    --in) in="$2"; shift 2;;
    --out) out="$2"; shift 2;;
    *) echo "Unknown arg: $1" >&2; exit 1;;
  esac
done

if [[ -z "$in" || -z "$out" ]]; then
  echo "Usage: $0 --in INPUT.rtf --out OUTPUT.rtf" >&2
  exit 1
fi

awk -v OUTFILE="$out" '
function hex2c(h) { return sprintf("%c", strtonum("0x" h)) }
function decode_hex(s,   a,b){
  # Match RTF hex escapes: \\'hh
  while (match(s, /\\\047[0-9A-Fa-f]{2}/)) {
    a = substr(s, 1, RSTART-1)
    b = substr(s, RSTART+2, 2)
    s = a hex2c(b) substr(s, RSTART+4)
  }
  return s
}
BEGIN {
  ORS = "";
  print "{\\rtf1\\ansi\\ansicpg1252\\deff0\n" > OUTFILE;
  print "{\\fonttbl{\\f0 Roboto;}}\n" >> OUTFILE;
  print "\\pard\\plain\\f0\\fs24\n" >> OUTFILE;
}
{
  gsub(/\r/, "");
  # process per record (paragraph)
  if ($0 ~ /\\listtext/) {
    # extract label
    if (match($0, /\\listtext[^}]*([0-9]+\.|[a-z]\.)\\tab/, m)) {
      label = m[1];
      body = $0;
      sub(/\\listtext[^}]*\}/, "", body);

      # decode \'hh hex escapes
      body = decode_hex(body);

      # drop \uNNNN? sequences (keep ASCII-only later)
      gsub(/\\u-?[0-9]+\??/, "", body);

      # remove control words
      gsub(/\\[A-Za-z]+-?[0-9]*[ ]?/, "", body);
      gsub(/\\[^A-Za-z]/, "", body);
      gsub(/[{}]/, "", body);

      # normalize spaces
      gsub(/[\n\t]/, "", body);
      gsub(/[ ]{2,}/, " ", body);
      sub(/^ +/, "", body);
      sub(/ +$/, "", body);

      # strip non-ascii
      gsub(/[^\x20-\x7E]/, "", body);

      if (length(body) > 0) {
        prefix = (label ~ /^[a-z]/ ? "\\tab " : "");
        text = label " " body;
        gsub(/[\\{}]/, "\\\\&", text);
        print prefix text "\\par\n" >> OUTFILE;
      }
    }
  }
}
END { print "}\n" >> OUTFILE; }
' RS="\\par" "$in"
