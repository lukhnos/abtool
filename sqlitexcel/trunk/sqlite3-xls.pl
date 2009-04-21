#!/usr/bin/perl -w -CS
use strict;
use utf8;
use DBI;
use Encode;
use Spreadsheet::WriteExcel;
use Getopt::Long;

my ($input, $output, $header, $cmd, $helpFlag, $verboseFlag);
my $optResult=GetOptions(
    "output=s" => \$output, "input=s" => \$input, "command=s" => \$cmd,
    "title=s" => \$header, "verbose" => \$verboseFlag, "help" => \$helpFlag);
                         
if ($helpFlag || (!$input || !$output || !$cmd)) {
    print "usage: sqlite3-xls -i input -o output -c command [-t header] [-v]\n";
    exit;
}
                         
my $workbook=Spreadsheet::WriteExcel->new($output);
my $worksheet=$workbook->add_worksheet();
my $dbh=DBI->connect("dbi:SQLite:dbname=$input", "", "");
my $sth=$dbh->prepare("$cmd");
$sth->execute();

my $rowcntr=0;
if ($header) {
    my @utf8data=map { decode_utf8($_); } split(/\s+/, $header);
    $worksheet->write($rowcntr++, 0, \@utf8data);
}

$|=1;
while(my @row=$sth->fetchrow_array) {
    my @utf8data=map { decode_utf8($_); } @row;
    if ($verboseFlag) { print "dumping $rowcntr\r"; }
    $worksheet->write($rowcntr++, 0, \@utf8data);
}
print "\n";
	
