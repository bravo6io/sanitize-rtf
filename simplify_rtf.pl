#!/usr/bin/env perl
use strict;
use warnings;

# Thin wrapper to avoid Perl parsing logic. Uses Node.js implementation.
my @args = @ARGV;
my $node = 'node';
my $script = 'simplify_rtf.js';

exec $node, $script, @args or die "Failed to run $node $script: $!\n";
