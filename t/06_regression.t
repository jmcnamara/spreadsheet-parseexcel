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

#use Test::More tests => 76;
use Test::More 'no_plan';

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
is( $got_1, $expected_1, $caption );

###############################################################################
#
# Test 3, 4. Unformatted cell value from a formatted cell.
#
$caption = "Test unformatted cell value from a formatted cell";
$cell = $worksheet->get_cell( 3, 1 );

$expected_1 = 40177;
$got_1      = $cell->unformatted();
$got_2      = $cell->{_Val};
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_1, $expected_1, $caption );

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
is( $got_1, $expected_1, $caption );

###############################################################################
#
# Test 7, 8. Cell format.
#
$caption = "Test cell format";
$cell = $worksheet->get_cell( 3, 1 );

$expected_1 = 172;
$got_1      = $cell->get_format()->{FmtIdx};
$got_2      = $cell->{Format}->{FmtIdx};;
$caption    = " \tCell regression: " . $caption;

is( $got_1, $expected_1, $caption );
is( $got_1, $expected_1, $caption );

###############################################################################
#
# Test 9. Cell format.
#
$caption = "Test cell format string";
$cell = $worksheet->get_cell( 2, 1 );

$expected_1 = '@'; # TODO. Probably should be '' or 'general'.
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

__END__
