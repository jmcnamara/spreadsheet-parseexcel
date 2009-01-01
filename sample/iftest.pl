use strict;
use warnings;

use Spreadsheet::ParseExcel;
die "Usage: $0 Excel/Test97.xls\n" if not @ARGV;
my $filename = $ARGV[0];
my $oBook = 
	Spreadsheet::ParseExcel::Workbook->Parse($filename);
my($iR, $iC, $oWkS, $oWkC);
foreach my $oWkS (@{$oBook->{Worksheet}}) {
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
#Sheet Name
print $oBook->Worksheet('Sheet1-ASC')->{Cells}[0][1]->Value, "\n";
#Sheet No
print $oBook->Worksheet(0)->{Cells}[0][1]->Value, "\n";
#Sheet Not found
print (($oBook->Worksheet('SHEET1') ? 'Exists' : 'Not Exists'), "\n");

__END__
# removed so we can run test (by Gabor)
use Spreadsheet::ParseExcel::SaveParser;
$oBook = 
	Spreadsheet::ParseExcel::SaveParser::Workbook->Parse($filename);
my $oWs = $oBook->AddWorksheet('TEST1');
$oWs->AddCell(10, 1, 'New Cell');
$oBook->SaveAs('iftest.xls');
