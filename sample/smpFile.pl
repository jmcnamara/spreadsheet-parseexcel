use lib qw(../);
use strict;
use Spreadsheet::ParseExcel;
my $oExcel = new Spreadsheet::ParseExcel;
sub PrnBook($);

#1.2 Normal Excel97
open(IN, 'Excel/Test97.xls');
binmode IN;
my $sWk;
read(IN, $sWk, 2000000);
close IN;
my $oBook1 = $oExcel->Parse([$sWk]);
PrnBook($oBook1);

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
