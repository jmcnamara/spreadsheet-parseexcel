use strict;
use warnings;

if(!(defined $ARGV[0])) {
    print<<EOF;
Usage: $0 Excel_File
EOF
    exit;
}
use Spreadsheet::ParseExcel;
main();

sub main {
my $oExcel = Spreadsheet::ParseExcel->new;
my $oBook = $oExcel->Parse($ARGV[0]);

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
            if($oWkC) {
                if($oWkC->{Rich}) {
                    foreach my $raR (@{$oWkC->{Rich}}) {
                        my $oFont = $raR->[1];
                        print
                            "--------------------------------------------\n",
                            'POS              :', $raR->[0], "\n",
                            'Name             :', $oFont->{Name}, "\n",
                            'Bold             :', $oFont->{Bold}, "\n",
                            'Italic           :', $oFont->{Italic}, "\n",
                            'Height           :', $oFont->{Height}, "\n",
                            'Underline        :', $oFont->{Underline}, "\n",
                            'UnderlineStyle   :', sprintf("%02x", $oFont->{UnderlineStyle}), "\n",
                            'Color            :', $oFont->{Color}, "\n",
                            'Color RGB        :', Spreadsheet::ParseExcel->ColorIdxToRGB($oFont->{Color}), "\n",
                            'Strikeout        :', $oFont->{Strikeout}, "\n",
                            'Super            :', $oFont->{Super}, "\n",
                    }
                }
            }
        }
    }
}
}
