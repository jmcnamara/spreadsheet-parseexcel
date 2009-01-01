#Khalid EZZARAOUI khalid@yromem.com
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
print '<?xml version="1.0" encoding="ISO-8859-1"?>', "\n";
print "<document>", "\n";

print "\t", "<meta>", "\n";
print "\t\t", "<file>", $oBook->{File} , "</file>", "\n";
print "\t\t", "<sheetcount>", $oBook->{SheetCount} , "</sheetcount>", "\n";
print "\t\t", "<author>", $oBook->{Author} , "</author>", "\n";
print "\t", "</meta>", "\n";

print "\t", "<sheets>", "\n";
#for(my $iSheet=0; $iSheet < $oBook->{SheetCount} ; $iSheet++) {
#    $oWkS = $oBook->{Worksheet}[$iSheet];
foreach my $oWkS (@{$oBook->{Worksheet}}) {
	print "\t\t", "<sheet " ;
	print "name=\"", $oWkS->{Name}, "\" >",  "\n";
	for(my $iR = $oWkS->{MinRow} ; defined $oWkS->{MaxRow} && $iR <= $oWkS->{MaxRow} ; $iR++) {
		print "\t\t\t", "<row num=\"", $iR, "\">", "\n" ;
		for(my $iC = $oWkS->{MinCol} ; defined $oWkS->{MaxCol} && $iC <= $oWkS->{MaxCol} ; $iC++) {
			$oWkC = $oWkS->{Cells}[$iR][$iC];
			print "\t\t\t\t", "<col num=\"", $iC, "\"", ">", $oWkC->Value, "</col>", "\n" if($oWkC);
			# print $oWkC->{_Kind}, "\n";
			}
		print "\t\t\t", "</row>", "\n" ;
	}
	print "\t\t", "</sheet>\n" ;
}
	
print "\t", "</sheets>", "\n";

print "</document>", "\n";

