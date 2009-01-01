use strict;
use warnings;

use Spreadsheet::ParseExcel;
my $oExcel = Spreadsheet::ParseExcel->new;
die "Usage: $0 file.xls\n" if not @ARGV;


#=Default
use Spreadsheet::ParseExcel::FmtDefault;
my $oFmt = Spreadsheet::ParseExcel::FmtDefault->new;
my $oBook = $oExcel->Parse($ARGV[0]);
#=cut

#Japan
#use Spreadsheet::ParseExcel::FmtJapan2;
#my $oFmt = Spreadsheet::ParseExcel::FmtJapan2->new(Code=>'sjis');
#my $oBook = $oExcel->Parse('Excel/FmtTest.xls', $oFmt);
#=cut

#Other Countries (ex. Russia (CP1251))
#use Spreadsheet::ParseExcel::FmtUnicode;
#my $oFmt = Spreadsheet::ParseExcel::FmtUnicode->new(Unicode_Map => 'CP1251');
#my $oBook = $oExcel->Parse('Excel/FmtTest.xls', $oFmt);
#=cut


my($iR, $iC, $oWkS, $oWkC);
print "=========================================\n";
print   'FILE             :', $oBook->{File} , "\n",
        'COUNT            :', $oBook->{SheetCount} , "\n",
        'AUTHOR           :', $oBook->{Author} , "\n";

for(my $iSheet=0; $iSheet < $oBook->{SheetCount} ; $iSheet++) {
    $oWkS = $oBook->{Worksheet}[$iSheet];
    print "--------- SHEET:", $oWkS->{Name}, "\n";

    print   ">> Print Setting\n",
            'Landscape        :', $oWkS->{Landscape}    , "\n",
            'Scale            :', $oWkS->{Scale}        , "\n",
            'FitWidth         :', $oWkS->{FitWidth}     , "\n",
            'FitHeight        :', $oWkS->{FitHeight}    , "\n",
            'PageFit          :', $oWkS->{PageFit}      , "\n",
            'PaperSize        :', $oWkS->{PaperSize}    , "\n",
            'PageStart        :', $oWkS->{PageStart}    , "\n",
            'UsePage          :', $oWkS->{UsePage}   , "\n";

    print   ">> Format\n",
            'Mergin Left      :', $oWkS->{LeftMergin}   , "\n",
            '       Right     :', $oWkS->{RightMergin}  , "\n",
            '       Top       :', $oWkS->{TopMergin}    , "\n",
            '       Bottom    :', $oWkS->{BottomMergin} , "\n",
            '       Header    :', $oWkS->{HeaderMergin} , "\n",
            '       Footer    :', $oWkS->{FooterMergin} , "\n",
            'Horizontal Center:', $oWkS->{HCenter}      , "\n",
            'Vertical Center  :', $oWkS->{VCenter}      , "\n",
            'Header           :', $oWkS->{Header}       , "\n",
            'Footer           :', $oWkS->{Footer}       , "\n";
    print   "Print Area       :\n";
    foreach my $raA (@{$oBook->{PrintArea}[$iSheet]}) {
        print '  Area            :', join(",", @$raA), "\n";
    }
    my $rhA = $oBook->{PrintTitle}[$iSheet];
    print "Print Title      :\n";
    print "          Row    :\n";
    foreach my $raTr (@{$rhA->{Row}}) {
        print '>>               :', join(",", @$raTr)   , "\n";
    }
    print "          Column :\n";
    foreach my $raTr (@{$rhA->{Column}}) {
        print '>>               :', join(",", @$raTr)   , "\n";
    }

    print   'Print Gridlines  :', $oWkS->{PrintGrid}    , "\n",
            'Print Headings   :', $oWkS->{PrintHeaders} , "\n",
            'NoColor          :', $oWkS->{NoColor}      , "\n",
            'Draft            :', $oWkS->{Draft}        , "\n",
            'Notes            :', $oWkS->{Notes}        , "\n",
            'LeftToRight      :', $oWkS->{LeftToRight}  , "\n";

    foreach my $raA (@{$oWkS->{MergedArea}}) {
        print "Merged Area:", join(",", @$raA), "\n";
    }
    print   'Horizontal PageBreak :', join(',', @{$oWkS->{HPageBreak}}), "\n" 
                            if($oWkS->{HPageBreak});
    print   'Vertical   PageBreak :', join(',', @{$oWkS->{VPageBreak}}), "\n"
                            if($oWkS->{VPageBreak});

    for(my $iR = $oWkS->{MinRow} ; 
            defined $oWkS->{MaxRow} && $iR <= $oWkS->{MaxRow} ; $iR++) {
        for(my $iC = $oWkS->{MinCol} ;
                        defined $oWkS->{MaxCol} && $iC <= $oWkS->{MaxCol} ; $iC++) {
            $oWkC = $oWkS->{Cells}[$iR][$iC];
            if($oWkC) {
                print   "------------------------------------------------------\n",
                        "( $iR , $iC ) =>", $oWkC->Value, "\n";
                print   'Format           :', $oFmt->FmtString($oWkC, $oBook), "\n",
                        'AlignH           :', $oWkC->{Format}->{AlignH}, "\n",
                        'AlignV           :', $oWkC->{Format}->{AlignV}, "\n",
                        'Indent           :', $oWkC->{Format}->{Indent}, "\n",
                        'Wrap             :', $oWkC->{Format}->{Wrap}, "\n",
                        'Shrink           :', $oWkC->{Format}->{Shrink}, "\n",
                        'Merged           :', (defined($oWkC->{Merged})? 
						$oWkC->{Merged}: 'No'), "\n",
                        'Rotate           :', $oWkC->{Format}->{Rotate}, "\n";
#                       'JustLast         :', $oWkC->{Format}->{JustLast}, "\n",
#                       'ReadDir          :', $oWkC->{Format}->{ReadDir}, "\n",

                my $oFont = $oWkC->{Format}->{Font};
                print   'Name             :', $oFont->{Name}, "\n",
                        'Bold             :', $oFont->{Bold}, "\n",
                        'Italic           :', $oFont->{Italic}, "\n",
                        'Height           :', $oFont->{Height}, "\n",
                        'Underline        :', $oFont->{Underline}, "\n",
                        'UnderlineStyle   :', sprintf("%02x", $oFont->{UnderlineStyle}), "\n",
                        'Color            :', $oFont->{Color}, "\n",
                        'Color RGB        :', Spreadsheet::ParseExcel->ColorIdxToRGB($oFont->{Color}), "\n",
                        'Strikeout        :', $oFont->{Strikeout}, "\n",
                        'Super            :', $oFont->{Super}, "\n",
                        'BdrStyle         :', join(',', @{$oWkC->{Format}->{BdrStyle}}), "\n",
                        'BdrColor         :', join(',', @{$oWkC->{Format}->{BdrColor}}), "\n",
                        'BdrDiag          :', join(',', @{$oWkC->{Format}->{BdrDiag}}), "\n",
                        'Pattern          :', join(',', @{$oWkC->{Format}->{Fill}}), "\n",
                        'Lock             :', $oWkC->{Format}->{Lock}, "\n",
                        'Hidden           :', $oWkC->{Format}->{Hidden}, "\n";
            }
        }
    }
}
