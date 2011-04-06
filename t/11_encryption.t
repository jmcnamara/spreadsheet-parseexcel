#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::ParseExcel.
#
# Tests for encrypted file handling.
#
# reverse('©'), April 2011, John McNamara, jmcnamara@cpan.org
#

use strict;

use Spreadsheet::ParseExcel;
use Test::More tests => 4;


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


###############################################################################
#
# Tests 1. Normal file, not encrypted.
#
$caption       = " \tUnencrypted file.";
$file          = 't/excel_files/chart1.xls';
$parser        = Spreadsheet::ParseExcel->new();
$workbook      = $parser->Parse( $file );
$error_string  = $parser->error();
$error_code    = $parser->error_code();
$expected_code = 0;

is( $error_code, $expected_code, $caption );


###############################################################################
#
# Tests 2. Encrypted file with default password.
#
$caption       = " \tEncrypted file. Defualt password.";
$file          = 't/excel_files/pers-protected.xls';
$parser        = Spreadsheet::ParseExcel->new();
$workbook      = $parser->Parse( $file );
$error_string  = $parser->error();
$error_code    = $parser->error_code();
$expected_code = 0;

is( $error_code, $expected_code, $caption );


###############################################################################
#
# Tests 3. Encrypted file with password.
#
$caption       = " \tEncrypted file with password.";
$file          = 't/excel_files/pers-encrypted-def-pass-QwErTyUiOp.xls';
$parser        = Spreadsheet::ParseExcel->new( Password => 'QwErTyUiOp' );
$workbook      = $parser->Parse( $file );
$error_string  = $parser->error();
$error_code    = $parser->error_code();
$expected_code = 0;

is( $error_code, $expected_code, $caption );


###############################################################################
#
# Tests 4. Encrypted file with password.
#
$caption       = " \tEncrypted file with password.";
$file          = 't/excel_files/pers-encrypted-RC4-pass-11.xls';
$parser        = Spreadsheet::ParseExcel->new( Password => '11' );
$workbook      = $parser->Parse( $file );
$error_string  = $parser->error();
$error_code    = $parser->error_code();
$expected_code = 3;

is( $error_code, $expected_code, $caption );


__END__
