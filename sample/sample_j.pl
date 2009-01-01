use strict;
use warnings;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtJapan;

if(($#ARGV< 0) or 
   (grep($ARGV[0] eq $_, ('euc', 'sjis', 'jis')) <= 0)) {
    print "USAGE: \n   > perl ", $0, " euc|sjis|jis\n";
    exit;
}
my $oExcel = Spreadsheet::ParseExcel->new;

#1. Make Formatter
my $oFmtJ = Spreadsheet::ParseExcel::FmtJapan->new(Code => $ARGV[0]);

#2.1 Test97
my $oBook = $oExcel->Parse('sample/Excel/Test97J.xls', $oFmtJ);
PrnBook($oBook);

#2.2 Test95
$oBook = $oExcel->Parse('sample/Excel/Test95J.xls', $oFmtJ);
PrnBook($oBook);

#2.3 1904 (1904 - 97)
$oBook = $oExcel->Parse('sample/Excel/Test1904.xls', $oFmtJ);
PrnBook($oBook);

#2.4 1904 (1904 - 95)
$oBook = $oExcel->Parse('sample/Excel/Test1904_95.xls', $oFmtJ);
PrnBook($oBook);

#2.5 1904 (1904 - 95)
$oBook = $oExcel->Parse('sample/Excel/AuthorK.xls', $oFmtJ);
PrnBook($oBook);

#2.6 1904 (1904 - 95)
$oBook = $oExcel->Parse('sample/Excel/AuthorK95.xls', $oFmtJ);
PrnBook($oBook);

sub PrnBook {
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
