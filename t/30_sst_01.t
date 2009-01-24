#!/usr/bin/perl

###############################################################################
#
# A test for Spreadsheet::ParseExcel.
#
# Tests for the SST with long strings over 2 CONTINUE blocks.
#
# reverse('©'), January 2009, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Spreadsheet::ParseExcel;
use Test::More tests => 40;

###############################################################################
#
# Tests setup.
#
my $parser;
my $workbook;
my $worksheet;
my $file;
my $cell;
my $row;
my $col;
my $got;
my $expected;
my $caption;
my $smiley          = chr 0x263a;
my $long_string_15k = ( 'x' x ( 15000 - 1 ) ) . 'z';
my $long_string_16k = ( 'x' x ( 16000 - 1 ) ) . 'z';
my $long_string_24k = ( 'x' x ( 24000 - 1 ) ) . 'z';
my $long_string_31k = ( 'x' x ( 31000 - 1 ) ) . 'z';

###############################################################################
###############################################################################
#
# File 1.
#
$file      = 't/excel_files/long_string1.xls';
$parser    = Spreadsheet::ParseExcel->new();
$workbook  = $parser->Parse($file);
$worksheet = $workbook->worksheet(0);

###############################################################################
#
# Test 1.
#
$row      = 0;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $smiley . $long_string_16k;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 2.
#
$row      = 1;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'Foo';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 3.
#
$row      = 2;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $smiley;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 4.
#
$row      = 3;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'This is a test string';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 5.
#
$row      = 4;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = "Smiley $smiley";
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
###############################################################################
#
# File 2.
#
$file      = 't/excel_files/long_string2.xls';
$parser    = Spreadsheet::ParseExcel->new();
$workbook  = $parser->Parse($file);
$worksheet = $workbook->worksheet(0);

###############################################################################
#
# Test 6.
#
$row      = 0;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $long_string_16k . $smiley;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 7.
#
$row      = 1;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'Foo';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 8.
#
$row      = 2;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $smiley;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 9.
#
$row      = 3;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'This is a test string';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 10.
#
$row      = 4;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = "Smiley $smiley";
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
###############################################################################
#
# File 3.
#
$file      = 't/excel_files/long_string3.xls';
$parser    = Spreadsheet::ParseExcel->new();
$workbook  = $parser->Parse($file);
$worksheet = $workbook->worksheet(0);

###############################################################################
#
# Test 11.
#
$row      = 0;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $smiley . $long_string_24k;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 12.
#
$row      = 1;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'Foo';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 13.
#
$row      = 2;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $smiley;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 14.
#
$row      = 3;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'This is a test string';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 15.
#
$row      = 4;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = "Smiley $smiley";
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
###############################################################################
#
# File 4.
#
$file      = 't/excel_files/long_string4.xls';
$parser    = Spreadsheet::ParseExcel->new();
$workbook  = $parser->Parse($file);
$worksheet = $workbook->worksheet(0);

###############################################################################
#
# Test 16.
#
$row      = 0;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $long_string_24k . $smiley;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 17.
#
$row      = 1;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'Foo';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 18.
#
$row      = 2;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $smiley;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 19.
#
$row      = 3;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'This is a test string';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 20.
#
$row      = 4;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = "Smiley $smiley";
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
###############################################################################
#
# File 5.
#
$file      = 't/excel_files/long_string5.xls';
$parser    = Spreadsheet::ParseExcel->new();
$workbook  = $parser->Parse($file);
$worksheet = $workbook->worksheet(0);

###############################################################################
#
# Test 21.
#
$row      = 0;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $smiley . $long_string_31k;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 22.
#
$row      = 1;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'Foo';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 23.
#
$row      = 2;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $smiley;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 24.
#
$row      = 3;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'This is a test string';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 25.
#
$row      = 4;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = "Smiley $smiley";
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
###############################################################################
#
# File 6.
#
$file      = 't/excel_files/long_string6.xls';
$parser    = Spreadsheet::ParseExcel->new();
$workbook  = $parser->Parse($file);
$worksheet = $workbook->worksheet(0);

###############################################################################
#
# Test 26.
#
$row      = 0;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $long_string_31k . $smiley;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 27.
#
$row      = 1;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'Foo';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 28.
#
$row      = 2;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $smiley;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 29.
#
$row      = 3;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'This is a test string';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 30.
#
$row      = 4;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = "Smiley $smiley";
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
###############################################################################
#
# File 7.
#
$file      = 't/excel_files/long_string7.xls';
$parser    = Spreadsheet::ParseExcel->new();
$workbook  = $parser->Parse($file);
$worksheet = $workbook->worksheet(0);

###############################################################################
#
# Test 31.
#
$row      = 0;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $long_string_15k . $smiley . $long_string_15k;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 32.
#
$row      = 1;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'Foo';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 33.
#
$row      = 2;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $smiley;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 34.
#
$row      = 3;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'This is a test string';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 35.
#
$row      = 4;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = "Smiley $smiley";
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
###############################################################################
#
# File 8.
#
$file      = 't/excel_files/long_string8.xls';
$parser    = Spreadsheet::ParseExcel->new();
$workbook  = $parser->Parse($file);
$worksheet = $workbook->worksheet(0);

###############################################################################
#
# Test 36.
#
$row      = 0;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $smiley . $long_string_15k . $smiley . $long_string_15k . $smiley;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 37.
#
$row      = 1;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'Foo';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 38.
#
$row      = 2;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = $smiley;
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 39.
#
$row      = 3;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = 'This is a test string';
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

###############################################################################
#
# Test 40.
#
$row      = 4;
$col      = 0;
$cell     = $worksheet->get_cell( $row, $col );
$expected = "Smiley $smiley";
$got      = $cell->value();
$caption  = " \tSST: File = $file, Row = $row, Col = $col";

is( $got, $expected, $caption );

