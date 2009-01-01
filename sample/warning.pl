#!/usr/bin/perl -w

use Spreadsheet::ParseExcel;

my $oBook = Spreadsheet::ParseExcel::Workbook->Parse('sample/Excel/gives-warnings.xls');
print "A1=" . $oBook->{Worksheet}->[0]->{Cells}[0][0]->Value . "\n";
print "A2=" . $oBook->{Worksheet}->[0]->{Cells}[1][0]->Value . "\n";

