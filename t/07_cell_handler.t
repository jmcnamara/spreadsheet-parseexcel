#!/usr/bin/perl

use strict;
use warnings;

BEGIN { delete @ENV{qw( LANG LC_ALL LC_DATE )}; }

use Test::More tests => 83;

use_ok ("Spreadsheet::ParseExcel");

my $file = "t/excel_files/Dates.xls";

my @expect = (# Date   Date      Date          Date          Text
    [ 39668,  "8-Aug", 20080808, "2008-08-08", "08/08/2008", "08 Aug 2008" ],
    [ 39672, "12-Aug", 20080812, "2008-08-12", "08/12/2008", "12 Aug 2008" ],
    [ 39790,  "8-Dec", 20081208, "2008-12-08", "12/08/2008", "08 Dec 2008" ],
    [ 39673, "13-Aug", 20080813, "2008-08-13", "08/13/2008", "13 Aug 2008" ],
    );

my $handler_number = 1;

my $cell_cnt;
my $handler1 = sub {
    my ($wb, $idx, $row, $col, $cell) = @_;
	my $R = $row + 1;
	my $C = $col + 1;
    return if $R > 4;
    cmp_ok($handler_number, '==', 1, 'Correct handler');
    parse_second_workbook() if ++$cell_cnt == 10;
	my $cell_pos = ("A".."E")[$col].$R;
	is ($cell->Value, $expect[$row][$C], "Handler 1 value for $cell_pos");
};

my $parser1;
ok (
	$parser1 = Spreadsheet::ParseExcel->new(
		CellHandler => $handler1,
		NotSetCell  => 1,
	),
	"Create parser 1",
);

my $handler2 = sub {
    my ($wb, $idx, $row, $col, $cell) = @_;
	my $R = $row + 1;
	my $C = $col + 1;
    return if $R > 4;
    cmp_ok($handler_number, '==', 2, 'Correct handler');
	my $cell_pos = ("A".."E")[$col].$R;
	is ($cell->Value, $expect[$row][$C], "Handler 2 value for $cell_pos");
};

my $parser2;
ok (
	$parser2 = Spreadsheet::ParseExcel->new(
		CellHandler => $handler2,
		NotSetCell  => 1,
	),
	"Create parser 2",
);

$parser1->parse($file);

sub parse_second_workbook {
  $handler_number = 2;
  $parser2->parse($file);
  $handler_number = 1;
}
