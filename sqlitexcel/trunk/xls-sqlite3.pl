#!/usr/bin/perl -CS

use strict;
use utf8;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtJapan;
use Encode;
use DBI;
use Getopt::Long;

my ($input, $output, $header, $cmd, $helpFlag, $verboseFlag, $sheet, $repeat,
    $upper);
$sheet=0;

my $optResult=GetOptions(
    "output=s" => \$output, "input=s" => \$input, "command=s" => \$cmd,
    "title" => \$header, "verbose" => \$verboseFlag, "help" => \$helpFlag,
    "sheet=s" => \$sheet, "repeat=s" => \$repeat, "uppercase" => \$upper);
                         
if ($helpFlag || (!$input || !$output || !$cmd)) {
    print "usage: xls-sqlite3 -i input -o output -c command [-s sheet] [-t] [-v] [-r column]\n";
    exit;
}

my $dbh=DBI->connect("dbi:SQLite:dbname=$output", "", "");
$dbh->do("begin;");
my $sth=$dbh->prepare($cmd);

my $workbook=Spreadsheet::ParseExcel::Workbook->Parse($input, 
    Spreadsheet::ParseExcel::FmtJapan->new);
my $worksheet=$workbook->{Worksheet}[int $sheet];
my $begin=$header ? 1 : 0;
$|=1;

for my $r ($begin..$worksheet->{MaxRow}) {
    if ($verboseFlag) { print "dumping row $r\r"; }
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
        if ($upper) {
            $v = toUpper($v);
        }
        push @data, $v;
    }
    if (defined $repeat) { my $x=$data[$repeat]; push @data, $x; }
#   print join(", ", @data), "\n";
    $sth->execute(@data);
}

$dbh->do("commit;");


sub toUpper {
    $_=shift;
    s/[éèëêÉÈËÊ]/E/g;
    s/[áàäâÁÀÄÂ]/A/g;
    s/[íìïîÍÌÏÎ]/I/g;
    s/[óòöôÓÒÖÔ]/O/g;
    s/[úùüûÚÙÜÔ]/U/g;
    s/[çÇ]/C/g;
    uc;
}