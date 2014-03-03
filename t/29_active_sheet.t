#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::ParseExcel.
#
# Test for get_active_sheet
#

use strict;

use Test::More tests => 8;

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

my $ws = $book->worksheet('Sheet1');
my $color = $book->color_idx_to_rgb($ws->get_tab_color);
is($color,'339966');

my $hidden = $ws->is_sheet_hidden;
is($hidden, 0);
$hidden = $book->worksheet('Sheet3')->is_sheet_hidden;
is($hidden, 1);

$hidden = $ws->is_row_hidden(1-1);
is($hidden,undef);
$hidden = $ws->is_row_hidden(4-1);
is($hidden,1);

$hidden = $ws->is_col_hidden(ord( 'A' ) - ord( 'A' ));
is($hidden,undef);
$hidden = $ws->is_col_hidden(ord( 'D' ) - ord( 'A' ));
is($hidden,1);




__END__
