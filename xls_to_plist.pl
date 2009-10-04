#!/usr/bin/perl

# Public domain. Written by Lukhnos D. Liu (lukhnos@lukhnos.org)
use strict;
# use open qw( :std :encoding(UTF-8) );
use utf8;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtJapan;
use Encode;
use Data::Plist::XMLWriter;

my $sheet = 0;
my $header = 0;

my $workbook = Spreadsheet::ParseExcel::Workbook->Parse($ARGV[0]);
my $worksheet = $workbook->{Worksheet}[int $sheet];
my $begin = $header ? 1 : 0;
$|=1;

my @output = ();

for my $r ($begin..$worksheet->{MaxRow}) {
    my @data=();
    for my $c (0..$worksheet->{MaxCol}) {
        my $cell=$worksheet->{Cells}[$r][$c];
        my $v=(defined $cell) ? $cell->Value : "";

		# no more need to determine if #cell->{Code} is UCS2-BE encoded
        if ($cell->{Code} eq "ucs2") {
            # $v=Encode::decode("UCS2-BE", $v);
        }
        else {            
            $v=Encode::encode("utf-8", $v);
        }
        push @data, $v;
    }
    
    push @output, \@data;
    # print join(", ", @data), "\n";
}

my $write = Data::Plist::XMLWriter->new;

print decode("utf8", $write->write(\@output));

