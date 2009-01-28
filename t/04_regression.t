#!/usr/bin/perl

###############################################################################
#
# A test for Spreadsheet::ParseExcel.
#
# Regression tests for Worksheet properties and methods.
#
# reverse('©'), January 2009, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::Utility 'sheetRef';
use Test::More 'no_plan';

###############################################################################
#
# Tests setup.
#
my $file     = 't/excel_files/worksheet_01.xls';
my $parser   = Spreadsheet::ParseExcel->new();
my $workbook = $parser->Parse($file);
my $worksheet;
my $cell;
my $got_1;
my $got_2;
my $expected;
my $caption;

###############################################################################
#
# Test 1.
#
$caption   = "Test string in cell A1";
$worksheet = $workbook->worksheet('Sheet1');
$cell      = $worksheet->get_cell( sheetRef('A1') );
$expected  = 'This is a test workbook for worksheet regression';
$got_1     = $cell->value();
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );

###############################################################################
#
# Test 2.
#
$caption   = "Test workshet name";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 'Sheet1';
$got_1     = $worksheet->{Name};
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );

###############################################################################
#
# Test 3, 4.
#
$caption   = "Test default row height";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 12.75;
$got_1     = $worksheet->{DefRowHeight};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 5, 6.
#
$caption   = "Test default column width";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 8.43;
$got_1     = $worksheet->{DefColWidth};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Tests 7, 8.
#
$caption   = "Test row '1' height";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 12.75;
$got_1     = $worksheet->{RowHeight}->[0];
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 9, 10.
#
$caption   = "Test column 'A' width";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = undef;
$got_1     = $worksheet->{ColWidth}->[0];
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 11, 12.
#
$caption   = "Test landscape print setting";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 1;
$got_1     = $worksheet->{Landscape};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 13, 14.
#
$caption   = "Test print scale";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 100;
$got_1     = $worksheet->{Scale};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 15, 16. Note, use Sheet3 for counter example.
#
$caption   = "Test print fit to page";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 0;
$got_1     = $worksheet->{PageFit};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 17, 18. Note, use Sheet3 for counter example.
#
$caption   = "Test print fit to page width";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 1;
$got_1     = $worksheet->{FitWidth};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 19, 20. Note, use Sheet3 for counter example.
#
$caption   = "Test print fit to page height";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 1;
$got_1     = $worksheet->{FitHeight};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 21, 22. Note, use Sheet3 for counter example.
#
$caption   = "Test paper size";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 9;
$got_1     = $worksheet->{PaperSize};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 23, 24.
#
$caption   = "Test user defined start page for printing";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 0;
$got_1     = $worksheet->{UsePage};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 25, 26.
#
$caption   = "Test user defined start page for printing";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 1;
$got_1     = $worksheet->{PageStart};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 27, 28.
#
$caption   = "Test left margin";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = undef;
$got_1     = $worksheet->{LeftMargin};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 29, 30.
#
$caption   = "Test right margin";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = undef;
$got_1     = $worksheet->{RightMargin};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 31, 32.
#
$caption   = "Test top margin";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = undef;
$got_1     = $worksheet->{TopMargin};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 33, 34.
#
$caption   = "Test bottom margin";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = undef;
$got_1     = $worksheet->{BottomMargin};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 35, 36.
#
$caption   = "Test header margin";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 0.5;
$got_1     = $worksheet->{HeaderMargin};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 37, 38.
#
$caption   = "Test footer margin";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 0.5;
$got_1     = $worksheet->{FooterMargin};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 39, 40.
#
$caption   = "Test center horizontally";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 0;
$got_1     = $worksheet->{HCenter};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 41, 42.
#
$caption   = "Test center vertically";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 0;
$got_1     = $worksheet->{VCenter};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 43, 44.
#
$caption   = "Test header";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = undef;
$got_1     = $worksheet->{Header};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 45, 46.
#
$caption   = "Test Footer";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = undef;
$got_1     = $worksheet->{Footer};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 47, 48.
#
$caption   = "Test print with gridlines";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 0;
$got_1     = $worksheet->{PrintGrid};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 49, 50.
#
$caption   = "Test print with row and column headers";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 0;
$got_1     = $worksheet->{PrintHeaders};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 51, 52.
#
$caption   = "Test print in black and white";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 0;
$got_1     = $worksheet->{NoColor};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 53, 54.
#
$caption   = "Test print in draft quality";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 0;
$got_1     = $worksheet->{Draft};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 55, 56.
#
$caption   = "Test print comments";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 0;
$got_1     = $worksheet->{Notes};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 57, 58.
#
$caption   = "Test print over then down";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 0;
$got_1     = $worksheet->{Notes};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 59, 60.
#
$caption   = "Test horizontal page breaks";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = undef;
$got_1     = $worksheet->{HPageBreak};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 61, 62.
#
$caption   = "Test vertical page breaks";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = undef;
$got_1     = $worksheet->{VPageBreak};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

###############################################################################
#
# Test 63, 64.
#
$caption   = "Test merged areas";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = undef;
$got_1     = $worksheet->{MergedArea};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

__END__
###############################################################################
#
# Test x, y.
#
$caption   = "Test xxx";
$worksheet = $workbook->worksheet('Sheet1');
$expected  = 'Sheet1';
$got_1     = $worksheet->{Name};
$got_2     = $got_1;
$caption   = " \tWorksheet regression: " . $caption;

is( $got_1, $expected, $caption );
is( $got_2, $expected, $caption );

