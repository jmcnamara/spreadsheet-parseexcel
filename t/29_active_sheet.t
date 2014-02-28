#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::ParseExcel.
#
# Test for get_active_sheet
#

use strict;

use Test::More tests => 1;

use Spreadsheet::ParseExcel;

##############################################################################
#
# Tests.
#

my $parser = Spreadsheet::ParseExcel->new;

# Workbook saved with sheet2 (index 1) open

my $book = $parser->parse( "t/excel_files/TestActiveSheet.xls" );
my $active = $book->get_active_sheet;

is($active, 1);


__END__
