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
use Test::More tests => 24;

###############################################################################
#
# Tests setup.
#
my $parser;
my $workbook;
my $worksheet;
my $file;
my $cell;
my $sheetname;
my $cell_value;
my $expected;
my $caption1;
my $caption2;

###############################################################################
#
# Chart 1.
#
$file     = 't/excel_files/chart1.xls';
$parser   = Spreadsheet::ParseExcel->new();
$workbook = $parser->Parse($file);

for my $index ( 0 .. 2 ) {
    $worksheet  = $workbook->worksheet($index);
    $sheetname  = $worksheet->{Name};
    $cell       = $worksheet->get_cell( 0, 0 );
    $cell_value = $cell->value();
    $expected   = 'Sheet' . ( $index + 1 );
    $caption1   = " \tFile = $file, Sheet name = $expected";
    $caption2   = " \t     + $file, Cell value = $expected";

    is( $sheetname,  $expected, $caption1 );
    is( $cell_value, $expected, $caption2 );
}

###############################################################################
#
# Chart 2.
#
$file     = 't/excel_files/chart2.xls';
$parser   = Spreadsheet::ParseExcel->new();
$workbook = $parser->Parse($file);

for my $index ( 0 .. 2 ) {
    $worksheet  = $workbook->worksheet($index);
    $sheetname  = $worksheet->{Name};
    $cell       = $worksheet->get_cell( 0, 0 );
    $cell_value = $cell->value();
    $expected   = 'Sheet' . ( $index + 1 );
    $caption1   = " \tFile = $file, Sheet name = $expected";
    $caption2   = " \t     + $file, Cell value = $expected";

    is( $sheetname,  $expected, $caption1 );
    is( $cell_value, $expected, $caption2 );
}

###############################################################################
#
# Chart 3.
#
$file     = 't/excel_files/chart3.xls';
$parser   = Spreadsheet::ParseExcel->new();
$workbook = $parser->Parse($file);

for my $index ( 0 .. 2 ) {
    $worksheet  = $workbook->worksheet($index);
    $sheetname  = $worksheet->{Name};
    $cell       = $worksheet->get_cell( 0, 0 );
    $cell_value = $cell->value();
    $expected   = 'Sheet' . ( $index + 1 );
    $caption1   = " \tFile = $file, Sheet name = $expected";
    $caption2   = " \t     + $file, Cell value = $expected";

    is( $sheetname,  $expected, $caption1 );
    is( $cell_value, $expected, $caption2 );
}

###############################################################################
#
# Chart 4.
#
$file     = 't/excel_files/chart4.xls';
$parser   = Spreadsheet::ParseExcel->new();
$workbook = $parser->Parse($file);

for my $index ( 0 .. 2 ) {
    $worksheet  = $workbook->worksheet($index);
    $sheetname  = $worksheet->{Name};
    $cell       = $worksheet->get_cell( 0, 0 );
    $cell_value = $cell->value();
    $expected   = 'Sheet' . ( $index + 1 );
    $caption1   = " \tFile = $file, Sheet name = $expected";
    $caption2   = " \t     + $file, Cell value = $expected";

    is( $sheetname,  $expected, $caption1 );
    is( $cell_value, $expected, $caption2 );
}

__END__
