use strict;
if(!(defined $ARGV[0])) {
    print<<EOF;
Usage: $0 Excel_File [Code]
  Code: euc, sjis, jis, ... (same as Jcode.pm)
EOF
    exit;
}
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtJapan;
my $oExcel = new Spreadsheet::ParseExcel;
my $oFmtJ = Spreadsheet::ParseExcel::FmtJapan->new(Code => $ARGV[1]);
my $oBook = $oExcel->Parse($ARGV[0], $oFmtJ);

my($iR, $iC, $oWkS, $oWkC);
print "=========================================\n";
print "FILE  :", $oBook->{File} , "\n";
print "COUNT :", $oBook->{SheetCount} , "\n";
print "AUTHOR:", $oBook->{Author} , "\n";
#for(my $iSheet=0; $iSheet < $oBook->{SheetCount} ; $iSheet++) {
#    $oWkS = $oBook->{Worksheet}[$iSheet];
for my $oWkS (@{$oBook->{Worksheet}}) {
    print "--------- SHEET:", $oWkS->{Name}, "\n";
    for(my $iR = $oWkS->{MinRow} ; 
            defined $oWkS->{MaxRow} && $iR <= $oWkS->{MaxRow} ; $iR++) {
        print "ROW HEIGHT:", $oWkS->{RowHeight}[$iR], "\n";
        for(my $iC = $oWkS->{MinCol} ;
                        defined $oWkS->{MaxCol} && $iC <= $oWkS->{MaxCol} ; $iC++) {
            $oWkC = $oWkS->{Cells}[$iR][$iC];
            print "( $iR , $iC ) =>", $oWkC->Value, "\n" if($oWkC);
        }
    }
}
