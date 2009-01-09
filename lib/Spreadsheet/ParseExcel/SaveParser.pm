# Spreadsheet::ParseExcel::SaveParser
#  by Kawai, Takanori (Hippo2000) 2001.5.1
# This Program is ALPHA version.
#//////////////////////////////////////////////////////////////////////////////
# Spreadsheet::ParseExcel:.SaveParser Objects
#//////////////////////////////////////////////////////////////////////////////

#==============================================================================
# Spreadsheet::ParseExcel::SaveParser
#==============================================================================
package Spreadsheet::ParseExcel::SaveParser;
use strict;
use warnings;

use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser::Workbook;
use Spreadsheet::ParseExcel::SaveParser::Worksheet;
use Spreadsheet::WriteExcel;
use base 'Spreadsheet::ParseExcel';
our $VERSION = '0.42';

use constant MagicCol => 1.14;

#------------------------------------------------------------------------------
# new (for Spreadsheet::ParseExcel::SaveParser)
#------------------------------------------------------------------------------
sub new {
    my ( $sPkg, %hKey ) = @_;
    $sPkg->SUPER::new(%hKey);
}

#------------------------------------------------------------------------------
# Create
#------------------------------------------------------------------------------
sub Create {
    my ( $oThis, $oWkFmt ) = @_;

    #0. New $oBook
    my $oBook = Spreadsheet::ParseExcel::Workbook->new;
    $oBook->{SheetCount} = 0;

    #2. Ready for format
    if ($oWkFmt) {
        $oBook->{FmtClass} = $oWkFmt;
    }
    else {
        $oBook->{FmtClass} = Spreadsheet::ParseExcel::FmtDefault->new;
    }
    return Spreadsheet::ParseExcel::SaveParser::Workbook->new($oBook);
}

#------------------------------------------------------------------------------
# Parse (for Spreadsheet::ParseExcel::SaveParser)
#------------------------------------------------------------------------------
sub Parse {
    my ( $oThis, $sFile, $oWkFmt ) = @_;
    my $oBook = $oThis->SUPER::Parse( $sFile, $oWkFmt );
    return undef unless ( defined $oBook );
    return Spreadsheet::ParseExcel::SaveParser::Workbook->new($oBook);
}

#------------------------------------------------------------------------------
# SaveAs (for Spreadsheet::ParseExcel::SaveParser)
#------------------------------------------------------------------------------
sub SaveAs {
    my ( $oThis, $oBook, $sName ) = @_;
    $oBook->SaveAs($sName);
}
1;

__END__

=head1 NAME

Spreadsheet::ParseExcel::SaveParser - Expand of Spreadsheet::ParseExcel with Spreadsheet::WriteExcel

=head1 SYNOPSIS

    #1. Write an Excel file with previous data
    use strict;
    use Spreadsheet::ParseExcel::SaveParser;
    my $oExcel = new Spreadsheet::ParseExcel::SaveParser;
    my $oBook = $oExcel->Parse('temp.xls');
    #1.1.Update and Insert Cells
    my $iFmt = $oBook->{Worksheet}[0]->{Cells}[0][0]->{FormatNo};
    $oBook->AddCell(0, 0, 0, 'No(UPD)',
        $oBook->{Worksheet}[0]->{Cells}[0][0]->{FormatNo});
    $oBook->AddCell(0, 1, 0, '304', $oBook->{Worksheet}[0]->{Cells}[0][0]);
    $oBook->AddCell(0, 1, 1, 'Kawai,Takanori', $iFmt);
    #1.2.add new worksheet
    my $iWkN = $oBook->AddWorksheet('Test');
    #1.3 Save
    $oExcel->SaveAs($oBook, 'temp.xls');  # as the same name
    $oExcel->SaveAs($oBook, 'temp1.xls'); # another name

    #2. Create new Excel file (most simple)
    use strict;
    use Spreadsheet::ParseExcel::SaveParser;
    my $oEx = new Spreadsheet::ParseExcel::SaveParser;
    my $oBook = $oEx->Create();
    $oBook->AddFormat;
    $oBook->AddWorksheet('NewWS');
    $oBook->AddCell(0, 0, 1, 'New Cell');
    $oEx->SaveAs($oBook, 'new.xls');

    #3. Create new Excel file(more complex)
    #!/usr/local/bin/perl
    use strict;
    use Spreadsheet::ParseExcel::SaveParser;
    my $oEx = new Spreadsheet::ParseExcel::SaveParser;
    my $oBook = $oEx->Create();
    my $iF1 = $oBook->AddFont(
            Name      => 'Arial',
            Height    => 11,
            Bold      => 1, #Bold
            Italic    => 1, #Italic
            Underline => 0,
            Strikeout => 0,
            Super     => 0,
        );
    my $iFmt =
    $oBook->AddFormat(
        Font => $oBook->{Font}[$iF1],
        Fill => [1, 10, 0],         # Filled with Red
                                    # cf. ParseExcel (@aColor)
        BdrStyle => [0, 1, 1, 0],   #Border Right, Top
        BdrColor => [0, 11, 0, 0],  # Right->Green
    );
    $oBook->AddWorksheet('NewWS');
    $oBook->AddCell(0, 0, 1, 'Cell', $iFmt);
    $oEx->SaveAs($oBook, 'new.xls');

I<new interface...>

    use strict;
    use Spreadsheet::ParseExcel::SaveParser;
    $oBook =
        Spreadsheet::ParseExcel::SaveParser::Workbook->Parse('Excel/Test97.xls');
    my $oWs = $oBook->AddWorksheet('TEST1');
    $oWs->AddCell(10, 1, 'New Cell');
    $oBook->SaveAs('iftest.xls');

=head1 DESCRIPTION

Spreadsheet::ParseExcel::SaveParser : Expand of Spreadsheet::ParseExcel with Spreadsheet::WriteExcel

=head2 Functions

=over 4

=item new

I<$oExcel> = new Spreadsheet::ParseExcel::SaveParser();

Constructor.

=item Parse

I<$oWorkbook> = $oParse->Parse(I<$sFileName> [, I<$oFmt>]);

return L<"Workbook"> object.
if error occurs, returns undef.

=over 4

=item I<$sFileName>

name of the file to parse (Same as Spreadsheet::ParseExcel)

From 0.12 (with OLE::Storage_Lite v.0.06),
scalar reference of file contents (ex. \$sBuff) or
IO::Handle object (inclucdng IO::File etc.) are also available.

=item I<$oFmt>

Formatter Class to format the value of cells.

=back

=item Create

I<$oWorkbook> = $oParse->Create([I<$oFmt>]);

return new L<"Workbook"> object.
if error occurs, returns undef.

=over 4

=item I<$oFmt>

Formatter Class to format the value of cells.

=back

=item SaveAs

I<$oWorkbook> = $oParse->SaveAs( $oBook, $sName);

save $oBook image as an Excel file named $sName.

=over 4

=item I<$oBook>

An Excel Workbook object to save.

=back

=item I<$sName>

Name of new Excel file.

=back

=head2 Workbook

I<Spreadsheet::ParseExcel::SaveParser::Workbook>

Workbook is a subclass of Spreadsheet::ParseExcel::Workbook.
And has these methods :

=over 4

=item AddWorksheet

I<$oWorksheet> = $oBook->AddWorksheet($sName, %hProperty);

Create new Worksheet(Spreadsheet::ParseExcel::Worksheet).

=over 4

=item I<$sName>

Name of new Worksheet

=item I<$hProperty>

Property of new Worksheet.

=back

=item AddFont

I<$oWorksheet> = $oBook->AddFont(%hProperty);

Create new Font(Spreadsheet::ParseExcel::Font).

=over 4

=item I<$hProperty>

Property of new Worksheet.

=back

=item AddFormat

I<$oWorksheet> = $oBook->AddFormat(%hProperty);

Create new Format(Spreadsheet::ParseExcel::Format).

=over 4

=item I<$hProperty>

Property of new Format.

=back

=item AddCell

I<$oWorksheet> = $oBook->AddCell($iWorksheet, $iRow, $iCol, $sVal, $iFormat [, $sCode]);

Create new Cell(Spreadsheet::ParseExcel::Cell).

=over 4

=item I<$iWorksheet>

Number of Worksheet

=back

=over 4

=item I<$iRow>

Number of row

=back

=over 4

=item I<$sVal>

Value of the cell.

=back

=over 4

=item I<$iFormat>

Number of format for use. To specify just same as another cell,
you can set it like below:

ex.

  $oCell=$oWorksheet->{Cells}[0][0]; #Just a sample
  $oBook->AddCell(0, 1, 0, 'New One', $oCell->{FormatNo});
    #or
  $oBook->AddCell(0, 1, 0, 'New One', $oCell);

=back

=over 4

=item I<$sCode>

  Character code

=back

=back

=head2 Worksheet

I<Spreadsheet::ParseExcel::SaveParser::Worksheet>

Worksheet is a subclass of Spreadsheet::ParseExcel::Worksheet.
And has these methods :

=over 4

=item AddCell

I<$oWorksheet> = $oWkSheet->AddCell($iRow, $iCol, $sVal, $iFormat [, $sCode]);

Create new Cell(Spreadsheet::ParseExcel::Cell).

=over 4

=item I<$iRow>

Number of row

=back

=over 4

=item I<$sVal>

Value of the cell.

=back

=over 4

=item I<$iFormat>

Number of format for use. To specify just same as another cell,
you can set it like below:

ex.

  $oCell=$oWorksheet->{Cells}[0][0]; #Just a sample
  $oWorksheet->AddCell(1, 0, 'New One', $oCell->{FormatNo});
    #or
  $oWorksheet->AddCell(1, 0, 'New One', $oCell);

=back

=over 4

=item I<$sCode>

  Character code

=back

=back

=head1 MORE INFORMATION

Please visit my Wiki page.
 I'll add sample at :
    http://www.hippo2000.info/cgi-bin/KbWikiE/KbWiki.pl

=head1 Known Problem

-Only last print area will remain. (Others will be removed)

=head1 AUTHOR

Kawai Takanori (Hippo2000) kwitknr@cpan.org

    http://member.nifty.ne.jp/hippo2000/            (Japanese)
    http://member.nifty.ne.jp/hippo2000/index_e.htm (English)

=head1 SEE ALSO

XLHTML, OLE::Storage, Spreadsheet::WriteExcel, OLE::Storage_Lite

This module is based on herbert within OLE::Storage and XLHTML.

=head1 COPYRIGHT

Copyright (c) 2000-2002 Kawai Takanori and Nippon-RAD Co. OP Division
All rights reserved.

You may distribute under the terms of either the GNU General Public
License or the Artistic License, as specified in the Perl README file.

=head1 ACKNOWLEDGEMENTS

First of all, I would like to acknowledge valuable program and modules :
XHTML, OLE::Storage and Spreadsheet::WriteExcel.

In no particular order: Yamaji Haruna, Simamoto Takesi, Noguchi Harumi,
Ikezawa Kazuhiro, Suwazono Shugo, Hirofumi Morisada, Michael Edwards, Kim Namusk
and many many people + Kawai Mikako.

=cut
