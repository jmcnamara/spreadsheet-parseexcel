#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::ParseExcel.
#
# Tests for parse() error handling.
#
# reverse('©'), August 2009, John McNamara, jmcnamara@cpan.org
#

use strict;

use Spreadsheet::ParseExcel;
use Test::More tests => 8;


###############################################################################
#
# Tests setup.
#
my $file;
my $parser;
my $workbook;
my $got;
my $caption;
my $error_code;
my $error_string;
my $expected_code;
my $expected_string;


###############################################################################
#
# Tests 1, 2. Normal file, no errors.
#
$caption         = " \tError codes: Normal file";
$file            = 't/excel_files/chart1.xls';
$parser          = Spreadsheet::ParseExcel->new();
$workbook        = $parser->Parse($file);
$error_string    = $parser->error();
$error_code      = $parser->error_code();
$expected_code   = 0;
$expected_string = '';

is( $error_code,   $expected_code,   $caption );
is( $error_string, $expected_string, $caption );


###############################################################################
#
# Tests 3, 4. Non existent file.
#
$caption         = " \tError codes: Non existent file";
$file            = 'file_doesnt_exist.xls';
$parser          = Spreadsheet::ParseExcel->new();
$workbook        = $parser->Parse($file);
$error_string    = $parser->error();
$error_code      = $parser->error_code();
$expected_code   = 1;
$expected_string = 'File not found';

is( $error_code,   $expected_code,   $caption );
is( $error_string, $expected_string, $caption );


###############################################################################
#
# Tests 5, 6. File with no Excel data.
#
$caption         = " \tError codes:File with no Excel data";
$file            = 't/00_basic.t';
$parser          = Spreadsheet::ParseExcel->new();
$workbook        = $parser->Parse($file);
$error_string    = $parser->error();
$error_code      = $parser->error_code();
$expected_code   = 2;
$expected_string = 'No Excel data found in file';

is( $error_code,   $expected_code,   $caption );
is( $error_string, $expected_string, $caption );


###############################################################################
#
# Tests 7, 8. Encrypted file.
#
$caption         = " \tEncrypted file";
$file            = 't/excel_files/encrypted.xls';
$parser          = Spreadsheet::ParseExcel->new();
$workbook        = $parser->Parse($file);
$error_string    = $parser->error();
$error_code      = $parser->error_code();
$expected_code   = 3;
$expected_string = 'File is encrypted';

is( $error_code,   $expected_code,   $caption );
is( $error_string, $expected_string, $caption );

__END__
