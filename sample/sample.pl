use strict;
use warnings;

use Spreadsheet::ParseExcel;
my $oExcel = Spreadsheet::ParseExcel->new;
sub PrnBook($);

#1.1 Normal Excel97
my $oBook = $oExcel->Parse('sample/Excel/Test97.xls');
PrnBook($oBook);

#1.2 Normal Excel95
$oBook = $oExcel->Parse('sample/Excel/Test95.xls');
PrnBook($oBook);

#1.3 Year 1904 Base (Excel97)
$oBook = $oExcel->Parse('sample/Excel/Test1904.xls');
PrnBook($oBook);

#1.4 Year 1904 Base (Excel95)
$oBook = $oExcel->Parse('sample/Excel/Test1904_95.xls');
PrnBook($oBook);

sub PrnBook($)
{
    my($oBook) = @_;
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
            for(my $iC = $oWkS->{MinCol} ;
                            defined $oWkS->{MaxCol} && $iC <= $oWkS->{MaxCol} ; $iC++) {
                $oWkC = $oWkS->{Cells}[$iR][$iC];
                print "( $iR , $iC ) =>", $oWkC->Value, "\n" if($oWkC);
            }
        }
    }
}
