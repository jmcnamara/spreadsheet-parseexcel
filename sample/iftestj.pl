use strict;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtJapan;
my $oBook = 
	Spreadsheet::ParseExcel::Workbook->Parse(
		'Excel/Test97J.xls',
		Spreadsheet::ParseExcel::FmtJapan->new (Code => 'sjis'));

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
