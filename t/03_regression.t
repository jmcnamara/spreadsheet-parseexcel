#!/usr/bin/perl

###############################################################################
#
# A test for Spreadsheet::ParseExcel.
#
# Regression tests for Workbook properties and methods.
#
# The testcases comprise tests direct hash access (old methodology) and method
# calls (>= version 0.50 methodology).
#
# reverse('©'), January 2009, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Spreadsheet::ParseExcel;

use Test::More tests => 16;

###############################################################################
#
# Tests setup.
#
my $file     = 't/excel_files/worksheet_01.xls';
my $parser   = Spreadsheet::ParseExcel->new();
my $workbook = $parser->Parse($file);
my @worksheets;
my $worksheet;
my $got;
my $expected;
my $caption;

###############################################################################
#
# Test 1 Test worksheets() method.
#
$caption    = "Test worksheet()";
@worksheets = $workbook->worksheets();
$got        = join ' ', map { $_->get_name } @worksheets;
$expected   = 'Sheet1 Sheet2 Sheet3';
$caption    = " \tWorkbook regression: " . $caption;

is( $got, $expected, $caption );

###############################################################################
#
# Test 2 Test deprecated Worksheets property.
#
$caption    = "Test worksheet()";
@worksheets = @{ $workbook->{Worksheet} };
$got        = join ' ', map { $_->get_name } @worksheets;
$expected   = 'Sheet1 Sheet2 Sheet3';
$caption    = " \tWorkbook regression: " . $caption;

is( $got, $expected, $caption );

###############################################################################
#
# Test 3 Test worksheet() method.
#
$caption   = "Test worksheet()";
$worksheet = $workbook->worksheet(1);
$got       = $worksheet->get_name();
$expected  = 'Sheet2';
$caption   = " \tWorkbook regression: " . $caption;

is( $got, $expected, $caption );

###############################################################################
#
# Test 4 Test worksheet() method.
#
$caption   = "Test worksheet()";
$worksheet = $workbook->worksheet('Sheet3');
$got       = $worksheet->get_name();
$expected  = 'Sheet3';
$caption   = " \tWorkbook regression: " . $caption;

is( $got, $expected, $caption );

###############################################################################
#
# Test 5 Test deprecated Worksheet() method
#
$caption   = "Test worksheet()";
$worksheet = $workbook->Worksheet('Sheet3');
$got       = $worksheet->get_name();
$expected  = 'Sheet3';
$caption   = " \tWorkbook regression: " . $caption;

is( $got, $expected, $caption );

###############################################################################
#
# Test 6 Test worksheet_count() method.
#
$caption  = "Test worksheet_count()";
$got      = $workbook->worksheet_count();
$expected = 3;
$caption  = " \tWorkbook regression: " . $caption;

is( $got, $expected, $caption );

###############################################################################
#
# Test 7 Test deprecated SheetCount property.
#
$caption  = "Test worksheet_count()";
$got      = $workbook->{SheetCount};
$expected = 3;
$caption  = " \tWorkbook regression: " . $caption;

is( $got, $expected, $caption );

###############################################################################
#
# Test 8 Test get_filename() method.
#
$caption  = "Test get_filename()";
$got      = $workbook->get_filename();
$expected = $file;
$caption  = " \tWorkbook regression: " . $caption;

is( $got, $expected, $caption );

###############################################################################
#
# Test 9 Test deprecated File property.
#
$caption  = "Test get_filename()";
$got      = $workbook->{File};
$expected = $file;
$caption  = " \tWorkbook regression: " . $caption;

is( $got, $expected, $caption );

###############################################################################
#
# Test 10 Test get_filename() method on a filehandle.
#
my $fh;
open $fh, '<', $file;
binmode $fh;
$workbook = $parser->Parse($fh);

$caption  = "Test get_filename()";
$got      = $workbook->get_filename();
$expected = undef;
$caption  = " \tWorkbook regression: " . $caption;

is( $got, $expected, $caption );

###############################################################################
#
# Test 11 Test get_print_areas() method.
#
$caption   = "Test get_print_areas()";
$worksheet = $workbook->worksheet(0);
$got       = $workbook->get_print_areas();
$expected =
  [ undef, [ [ 0, 0, 14, 2 ] ], [ [ 0, 1, 65535, 1 ], [ 3, 4, 7, 6 ] ] ];
$caption = " \tWorkbook regression: " . $caption;

is_deeply( $got, $expected, $caption );

###############################################################################
#
# Test 12 Test deprecated PrintArea property.
#
$caption   = "Test get_print_areas()";
$worksheet = $workbook->worksheet(0);
$got       = $workbook->{PrintArea};
$expected =
  [ undef, [ [ 0, 0, 14, 2 ] ], [ [ 0, 1, 65535, 1 ], [ 3, 4, 7, 6 ] ] ];
$caption = " \tWorkbook regression: " . $caption;

is_deeply( $got, $expected, $caption );

###############################################################################
#
# Test 13 Test get_print_titles() method.
#
$caption   = "Test get_print_titles()";
$worksheet = $workbook->worksheet(0);
$got       = $workbook->get_print_titles();
$expected  = [
    undef,
    {
        'Row'    => [ [ 0, 1 ] ],
        'Column' => [ [ 0, 2 ] ],

    },
    {
        'Row'    => [],
        'Column' => [ [ 0, 7 ] ],
    }
];
$caption = " \tWorkbook regression: " . $caption;

is_deeply( $got, $expected, $caption );

###############################################################################
#
# Test 14 Test deprecated PrintArea property.
#
$caption   = "Test get_print_titles()";
$worksheet = $workbook->worksheet(0);
$got       = $workbook->{PrintTitle};
$expected  = [
    undef,
    {
        'Row'    => [ [ 0, 1 ] ],
        'Column' => [ [ 0, 2 ] ],

    },
    {
        'Row'    => [],
        'Column' => [ [ 0, 7 ] ],
    }
];
$caption = " \tWorkbook regression: " . $caption;

is_deeply( $got, $expected, $caption );

###############################################################################
#
# Test 15 Test using_1904_date() method.
#
$caption  = "Test using_1904_date()";
$got      = $workbook->using_1904_date();
$expected = 0;
$caption  = " \tWorkbook regression: " . $caption;

is( $got, $expected, $caption );

###############################################################################
#
# Test 16 Test using_1904_date() method.
#
$workbook = $parser->Parse('t/excel_files/Dates1904.xls');    # File change.
$caption  = "Test using_1904_date()";
$got      = $workbook->using_1904_date();
$expected = 1;
$caption  = " \tWorkbook regression: " . $caption;

is( $got, $expected, $caption );

__END__

