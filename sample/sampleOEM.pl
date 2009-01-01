use strict;
use warnings;

use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtJapan;
use Spreadsheet::ParseExcel::FmtJapan2;

my $oExcel = Spreadsheet::ParseExcel->new;
my $oFmtJ = Spreadsheet::ParseExcel::FmtJapan->new(Code => 'euc');
my $oFmtJ2 = Spreadsheet::ParseExcel::FmtJapan2->new(Code => 'euc');
my $oBook = $oExcel->Parse('Excel/oem.xls', $oFmtJ);
PrnAll($oBook);
$oBook = $oExcel->Parse('Excel/oem.xls', $oFmtJ2);
PrnAll($oBook);

sub PrnAll {
    my ($oBook) = @_;
    my($iR, $iC, $oWkS, $oWkC);
    print "=========================================\n";
    print "FILE  :", $oBook->{File} , "\n";
    print "COUNT :", $oBook->{SheetCount} , "\n";
    print "AUTHOR:", $oBook->{Author} , "\n";
    for(my $iSheet=0; $iSheet < $oBook->{SheetCount} ; $iSheet++) {
        $oWkS = $oBook->{Worksheet}[$iSheet];
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
}
