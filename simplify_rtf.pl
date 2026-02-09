#!/usr/bin/env perl
#
# simplify_rtf.pl
# Usage:
#   perl simplify_rtf.pl --in INPUT.rtf --out OUTPUT.rtf
#
# What it does:
# - Reads a Word-generated RTF file
# - Extracts list items from \listtext entries
# - Writes a minimal RTF using Roboto 12pt with:
#   - Level 1: "1. 2. 3. ..."
#   - Level 2: "a. b. c. ..." (indented with \tab)

use strict;
use warnings;
use Getopt::Long qw(GetOptions);

my $in = '';
my $out = '';
GetOptions(
  'in=s'  => \$in,
  'out=s' => \$out,
) or die "Usage: $0 --in INPUT.rtf --out OUTPUT.rtf\n";

if (!$in || !$out) {
  die "Usage: $0 --in INPUT.rtf --out OUTPUT.rtf\n";
}

open my $fh, '<', $in or die "Failed to open input: $in\n";
local $/;
my $s = <$fh>;
close $fh;

$s =~ s/\r\n/\n/g;

sub rtf_unicode {
  my ($n) = @_;
  $n += 65536 if $n < 0;
  # Map common “smart” punctuation to ASCII equivalents.
  return '"' if $n == 8220 || $n == 8221;
  return "'" if $n == 8216 || $n == 8217;
  return '-' if $n == 8211 || $n == 8212;
  return chr($n) if $n >= 0 && $n < 128;
  return '';
}

open my $outfh, '>', $out or die "Failed to open output: $out\n";

print $outfh "{\\rtf1\\ansi\\ansicpg1252\\deff0\n";
print $outfh "{\\fonttbl{\\f0 Roboto;}}\n";
print $outfh "\\pard\\plain\\f0\\fs24\n";

while ($s =~ /\\listtext.*?([0-9]+\.|[a-z]\.)\\tab\}?(.+?)(?=\\par(?![a-z]))/sig) {
  my ($label, $body) = ($1, $2);
  # Decode and strip RTF control codes to plain ASCII text.
  $body =~ s/\\u(-?\d+)\??/rtf_unicode($1)/ge;
  $body =~ s/\\\'([0-9a-fA-F]{2})/chr(hex($1))/ge;
  $body =~ s/[\n\r\t]+//g;
  $body =~ s/\\[a-zA-Z]+-?\d* ?//g;
  $body =~ s/\\[^a-zA-Z]//g;
  $body =~ s/[{}]//g;
  $body =~ s/[ ]{2,}/ /g;
  $body =~ s/^\s+|\s+$//g;
  next unless length $body;

  # Level 2 items are indented with a tab.
  my $prefix = ($label =~ /^[a-z]/i) ? "\\tab " : "";
  my $text = $label . " " . $body;
  $text =~ s/([\\{}])/\\$1/g;
  $text =~ s/[^\x00-\x7F]//g;
  print $outfh $prefix, $text, "\\par\n";
}

print $outfh "}\n";
close $outfh;
