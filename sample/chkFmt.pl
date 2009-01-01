use strict;
use warnings;
use Spreadsheet::ParseExcel::Utility;
my $sWk;
my @aData = (-2, 0, 1234, -1234, 34323.23232, -233232.3233, 
        Spreadsheet::ParseExcel::Utility::LocaltimeExcel(13, 11, 12, 23, 2, 64, undef, 238));
my %hFmtTest = (
    0x00 => '@',
    0x01 => '0',
    0x02 => '0.00',
    0x03 => '#,##0',
    0x04 => '#,##0.00',
    0x05 => '(\\\\#,##0_);(\\\\#,##0)',
    0x06 => '(\\\\#,##0_);[RED](\\\\#,##0)',
    0x07 => '(\\\\#,##0.00_);(\\\\#,##0.00_)',
    0x08 => '(\\\\#,##0.00_);[RED](\\\\#,##0.00_)',
    0x09 => '0%',
    0x0A => '0.00%',
    0x0B => '0.00E+00',
    0x0C => '# ?/?',
    0x0D => '# ??/??',
    0x0E => 'm/d/yy',
    0x0F => 'd-mmm-yy',
    0x10 => 'd-mmm',
    0x11 => 'mmm-yy',
    0x12 => 'h:mm AM/PM',
    0x13 => 'h:mm:ss AM/PM',
    0x14 => 'h:mm',
    0x15 => 'h:mm:ss',
    0x16 => 'm/d/yy h:mm',
#0x17-0x24 -- Differs in Natinal
    0x25 => '(#,##0_);(#,##0)',
    0x26 => '(#,##0_);[RED](#,##0)',
    0x27 => '(#,##0.00);(#,##0.00)',
    0x28 => '(#,##0.00);[RED](#,##0.00)',
    0x29 => '_(*#,##0_);_(*(#,##0);_(*"-"_);_(@_)',
    0x2A => '_(\\\\*#,##0_);_(\\\\*(#,##0);_(*"-"_);_(@_)',
    0x2B => '_(*#,##0.00_);_(*(#,##0.00);_(*"-"??_);_(@_)',
    0x2C => '_(\\\\*#,##0.00_);_(\\\\*(#,##0.00);_(*"-"??_);_(@_)',
    0x2D => 'mm:ss',
    0x2E => '[h]:mm:ss',
    0x2F => 'mm:ss.0',
    0x30 => '##0.0E+0',
    0x31 => '@',
);
foreach my $sKey (sort {$a <=> $b} keys(%hFmtTest)) {
    my $sVal = $hFmtTest{$sKey};
    printf "============ %02x \n", $sKey;
    foreach my $sDt (@aData) {
        $sWk = Spreadsheet::ParseExcel::Utility::ExcelFmt($sVal, $sDt);
        printf "Fmt: %-20s : %-10s: Data: %s\n", $sVal, "$sDt", $sWk;
    }
}
