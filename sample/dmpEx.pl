use strict;
if(!(defined $ARGV[0])) {
    print<<EOF;
Usage: $0 Excel_File
EOF
    exit;
}
use Spreadsheet::ParseExcel;
my $oExcel = new Spreadsheet::ParseExcel;
my $oBook = $oExcel->Parse($ARGV[0]);

my($iR, $iC, $oWkS, $oWkC);
print "=========================================\n";
print "FILE  :", $oBook->{File} , "\n";
print "COUNT :", $oBook->{SheetCount} , "\n";
print "AUTHOR:", $oBook->{Author} , "\n";
#for(my $iSheet=0; $iSheet < $oBook->{SheetCount} ; $iSheet++) {
#    $oWkS = $oBook->{Worksheet}[$iSheet];
foreach my $oWkS (@{$oBook->{Worksheet}}) {
    print "--------- SHEET:", $oWkS->{Name}, "\n";
    for(my $iR = $oWkS->{MinRow} ; 
            defined $oWkS->{MaxRow} && $iR <= $oWkS->{MaxRow} ; $iR++) {
        for(my $iC = $oWkS->{MinCol} ;
                        defined $oWkS->{MaxCol} && $iC <= $oWkS->{MaxCol} ; $iC++) {
            $oWkC = $oWkS->{Cells}[$iR][$iC];
            print "( $iR , $iC ) =>", $oWkC->Value, "\n" if($oWkC);
print $oWkC->{_Kind}, "\n";
        }
    }
}
