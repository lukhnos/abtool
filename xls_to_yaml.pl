#!/usr/bin/env perl -CS

# Public domain. Written by Lukhnos D. Liu (lukhnos@lukhnos.org)

use strict;
use utf8;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtJapan;
use Encode;
use YAML;

my $sheet = 0;
my $header = 0;

my $workbook = Spreadsheet::ParseExcel::Workbook->Parse($ARGV[0], Spreadsheet::ParseExcel::FmtJapan->new);
my $worksheet = $workbook->{Worksheet}[int $sheet];
my $begin = $header ? 1 : 0;
$|=1;

my @output = ();

for my $r ($begin..$worksheet->{MaxRow}) {
    my @data=();
    for my $c (0..$worksheet->{MaxCol}) {
        my $cell=$worksheet->{Cells}[$r][$c];
        my $v=(defined $cell) ? $cell->Value : "";
        if ($cell->{Code} eq "ucs2") {
            $v=Encode::decode("UCS2-BE", $v);
        }
        else {
            $v=Encode::decode("iso8859-1", $v);
        }

        push @data, $v;
    }
    
    push @output, \@data;
    # print join(", ", @data), "\n";
}


print YAML::Dump(\@output);