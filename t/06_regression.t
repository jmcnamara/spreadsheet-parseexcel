#!/usr/bin/perl

###############################################################################
#
# A test for Spreadsheet::ParseExcel.
#
# Regression tests for Cell properties and methods.
#
# The tests are mainly in pairs where direct hash access (old methodology)
# is tested along with the method calls (>= version 0.50 methodology).
#
# reverse('©'), January 2009, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Spreadsheet::ParseExcel;
use Test::More tests => 32;

###############################################################################
#
# Tests setup.
#
my $file      = 't/excel_files/worksheet_01.xls';
my $parser    = Spreadsheet::ParseExcel->new();
my $workbook  = $parser->Parse($file);
my $worksheet = $workbook->worksheet('Sheet3');
my $cell;
my $format;
my $got_1;
my $got_2;
my $expected_1;
my $expected_2;
my $caption;

###############################################################################
#
# Test 1, 2. Unformatted cell value.
#
$caption = "Test unformatted cell value";
$cell = $worksheet->get_cell( 2, 1 );

$expected_1 = 1;
$got_1      = $cell->value();
$got_2      = $cell->Value();
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_1, $caption );

###############################################################################
#
# Test 3, 4. Unformatted cell value from a formatted cell.
#
$caption = "Test unformatted cell value from a formatted cell";
$cell = $worksheet->get_cell( 3, 1 );

$expected_1 = 40177;
$got_1      = $cell->unformatted();
$got_2      = $cell->{Val};
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_1, $caption );

###############################################################################
#
# Test 5, 6. Formatted cell value.
#
$caption = "Test formatted cell value";
$cell = $worksheet->get_cell( 3, 1 );

$expected_1 = '2009/12/30';
$got_1      = $cell->value();
$got_2      = $cell->Value;
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_1, $caption );

###############################################################################
#
# Test 7, 8. Cell format.
#
$caption = "Test cell format";
$cell = $worksheet->get_cell( 3, 1 );

$expected_1 = 170;
$got_1      = $cell->get_format()->{FmtIdx};
$got_2      = $cell->{Format}->{FmtIdx};;
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_1, $caption );

###############################################################################
#
# Test 9. Cell format.
#
$caption = "Test cell format string";
$cell = $worksheet->get_cell( 2, 1 );

$expected_1 = 'General'; # TODO. Probably should be '' or 'general'.
$got_1      = $workbook->{FmtClass}->FmtString( $cell, $workbook );
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );

###############################################################################
#
# Test 10. Cell format.
#
$caption = "Test cell format string";
$cell = $worksheet->get_cell( 3, 1 );

$expected_1 = 'yyyy/mm/dd';
$got_1      = $workbook->{FmtClass}->FmtString( $cell, $workbook );
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );

###############################################################################
#
# Test 11, 12. Cell type.
#
$caption = "Test cell type = Text";
$cell = $worksheet->get_cell( 2, 0 );

$expected_1 = 'Text';
$got_1      = $cell->type();
$got_2      = $cell->{Type};
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_1, $caption );


###############################################################################
#
# Test 13, 14. Cell type.
#
$caption = "Test cell type = Numeric";
$cell = $worksheet->get_cell( 2, 1 );

$expected_1 = 'Numeric';
$got_1      = $cell->type();
$got_2      = $cell->{Type};
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_1, $caption );

###############################################################################
#
# Test 15, 16. Cell type.
#
$caption = "Test cell type = Date";
$cell = $worksheet->get_cell( 3, 1 );

$expected_1 = 'Date';
$got_1      = $cell->type();
$got_2      = $cell->{Type};
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_1, $caption );


###############################################################################
#
# Test 17, 18. Cell string encoding.
#
$caption = "Test cell encoding = Ascii";
$cell = $worksheet->get_cell( 5, 0 );

$expected_1 = 1;
$expected_2 = undef;
$got_1      = $cell->encoding();
$got_2      = $cell->{Code};
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_2, $caption );


###############################################################################
#
# Test 19, 20. Cell string encoding.
#
$caption = "Test cell encoding = Unicode";
$cell = $worksheet->get_cell( 5, 1 );

$expected_1 = 2;
$expected_2 = 'ucs2';
$got_1      = $cell->encoding();
$got_2      = $cell->{Code};
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_2, $caption );


###############################################################################
#
# Test 21, 22. Cell string encoding.
#

# Switch file to get native encoding from Excel 5 file.
$file      = 't/excel_files/Test95J.xls';
$parser    = Spreadsheet::ParseExcel->new();
$workbook  = $parser->Parse($file);
$worksheet = $workbook->worksheet(0);

$caption = "Test cell encoding = Native";
$cell = $worksheet->get_cell( 1, 0 );

$expected_1 = 3;
$expected_2 = '_native_';
$got_1      = $cell->encoding();
$got_2      = $cell->{Code};
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_2, $caption );

# Switch back to main test worksheet.
$file      = 't/excel_files/worksheet_01.xls';
$parser    = Spreadsheet::ParseExcel->new();
$workbook  = $parser->Parse($file);
$worksheet = $workbook->worksheet('Sheet3');


###############################################################################
#
# Test 23, 24. Cell is un-merged.
#
$caption = "Test cell is un-merged";
$cell = $worksheet->get_cell( 6, 0 );

$expected_1 = undef;
$got_1      = $cell->is_merged();
$got_2      = $cell->{Merged};
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_1, $caption );


###############################################################################
#
# Test 25, 26. Cell is merged.
#
$caption = "Test cell is merged";
$cell = $worksheet->get_cell( 7, 0 );

$expected_1 = 1;
$got_1      = $cell->is_merged();
$got_2      = $cell->{Merged};
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_1, $caption );


###############################################################################
#
# Test 27, 28. Cell is merged.
#
$caption = "Test cell is merged";
$cell = $worksheet->get_cell( 7, 1 );

$expected_1 = 1;
$got_1      = $cell->is_merged();
$got_2      = $cell->{Merged};
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_1, $caption );


###############################################################################
#
# Test 29, 30. Cell is merged.
#
$caption = "Test cell is merged";
$cell = $worksheet->get_cell( 7, 1 );

$expected_1 = 1;
$got_1      = $cell->is_merged();
$got_2      = $cell->{Merged};
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_2, $expected_1, $caption );


###############################################################################
#
# Test 31, 32. Cell contains Rich Text.
#
$caption = "Test cell containing rich text";
$cell = $worksheet->get_cell( 8, 0 );

$expected_1 = [ 10, 14, 19 ];
$got_1      = $cell->get_rich_text();
$got_2      = $cell->{Rich};
$caption    = " \tCell regression: " . $caption;

# Just test the indices and not the font objects.
$got_1 = [ map { $_->[0] } @$got_1 ];
$got_2 = [ map { $_->[0] } @$got_2 ];

is_deeply( $got_1, $expected_1, $caption );
is_deeply( $got_2, $expected_1, $caption );


__END__
