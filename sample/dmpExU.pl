# This script requieres Unicode::Map
use strict;
if($#ARGV < 1) {
    print<<EOF;
Usage: $0 Excel_File [Code]
  Code:  CP932, CP1251, ... (same as Unicode::Map)
EOF
    exit;
}
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtUnicode;
my $oExcel = new Spreadsheet::ParseExcel;
my $oFmtJ = Spreadsheet::ParseExcel::FmtUnicode->new(Unicode_Map => $ARGV[1]);
my $oBook = $oExcel->Parse($ARGV[0], $oFmtJ);

my($iR, $iC, $oWkS, $oWkC);
print "=========================================\n";
print "FILE  :", $oBook->{File} , "\n";
print "COUNT :", $oBook->{SheetCount} , "\n";
print "AUTHOR:", $oBook->{Author} , "\n";

my $table = [];

for(my $iSheet=0; $iSheet < $oBook->{SheetCount} ; $iSheet++) {
    $oWkS = $oBook->{Worksheet}[$iSheet];
    print "--------- SHEET:", $oWkS->{Name}, "\n";
    for(my $iR = $oWkS->{MinRow} ; 
            defined $oWkS->{MaxRow} && $iR <= $oWkS->{MaxRow} ; $iR++) {
#       print "ROW HEIGHT:", $oWkS->{RowHeight}[$iR], "\n";
        for(my $iC = $oWkS->{MinCol} ;
                        defined $oWkS->{MaxCol} && $iC <= $oWkS->{MaxCol} ; $iC++) {
            $oWkC = $oWkS->{Cells}[$iR][$iC];
           print "( $iR , $iC ) =>", $oWkC->Value, "\n" if($oWkC);
        }
    }
}
