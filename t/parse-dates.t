#!/usr/bin/perl

use strict;
use warnings;

BEGIN { delete @ENV{qw( LANG LC_ALL LC_DATE )}; }

use Test::More tests => 151;

use_ok ("Spreadsheet::ParseExcel");

my $bk;
ok ($bk = Spreadsheet::ParseExcel::Workbook->Parse ("sample/Excel/Dates.xls"),
    "Open/parse Dates.xls");

ok (my @ws = @{$bk->{Worksheet}},	"Book has sheets");
is (scalar @ws, 1,			"But just one");

my $ws = $ws[0];
ok (my $cells = $ws->{Cells},		"Worksheet has cells");
is ($ws->{Name}, "DateTest",		"Worksheet label");
is ($ws->{MinRow}, 0,			"Row start");
is ($ws->{MinCol}, 0,			"Col start");
is ($ws->{MaxRow}, 6,			"Row count");
is ($ws->{MaxCol}, 4,			"Col count");

my @expect = (# Date   Date      Date          Date          Text
    [ 39668,  "8-Aug", 20080808, "2008-08-08", "08/08/2008", "08 Aug 2008" ],
    [ 39672, "12-Aug", 20080812, "2008-08-12", "08/12/2008", "12 Aug 2008" ],
    [ 39790,  "8-Dec", 20081208, "2008-12-08", "12/08/2008", "08 Dec 2008" ],
    [ 39673, "13-Aug", 20080813, "2008-08-13", "08/13/2008", "13 Aug 2008" ],
    );
#            non given, ISO,     type 14,      US broken,    Text-default
my @format = (undef,  "yyyymmdd", undef,       "mm/dd/yyyy", undef);
my @col = ("A".."E");

foreach my $row (0 .. 3) {
    foreach my $col (0 .. 4) {
	my $R = $row + 1;
	my $C = $col + 1;
	my $cell = ("A".."E")[$col].$R;
	ok (my $wc = $ws->{Cells}[$row][$col], "Cell $cell");
	is ($wc->{Val}, $expect[$row][$col == 4 ? 5 : 0], "Base value for $cell");
	is ($wc->Value, $expect[$row][$C], "Formatted value for $cell");

	is ($wc->{Type}, $col == 4 ? "Text" : "Date", "Cell $cell Type");
	my $fmt    = $wc->{Format};
	my $fmtstr = $bk->{FormatStr}{$fmt->{FmtIdx}};
	$fmtstr and $fmtstr =~ s/\\//g;
	is ($fmtstr, $format[$col], "Format string");
	}
    }

# Additional allignment tests
is ($ws->{Cells}[0][$_]{Format}{AlignV}, 3, "$col[$_]1 v-aligned justified") for 1..3;
is ($ws->{Cells}[1][$_]{Format}{AlignV}, 0, "$col[$_]2 v-aligned top")       for 1..3;
is ($ws->{Cells}[2][$_]{Format}{AlignV}, 1, "$col[$_]3 v-aligned center")    for 1..3;
is ($ws->{Cells}[3][$_]{Format}{AlignV}, 2, "$col[$_]4 v-aligned bottom")    for 1..3;

is ($ws->{Cells}[$_][0]{Format}{AlignH}, 0, "A$_ h-aligned -")       for 1..3;
is ($ws->{Cells}[$_][1]{Format}{AlignH}, 1, "B$_ h-aligned left")    for 1..3;
is ($ws->{Cells}[$_][2]{Format}{AlignH}, 2, "C$_ h-aligned center")  for 1..3;
is ($ws->{Cells}[$_][3]{Format}{AlignH}, 3, "D$_ h-aligned right")   for 1..3;

# Additional color tests
is ($ws->{Cells}[0][$_]{Format}{Font}{Color}, 32767, "$col[$_]1 font color") for 0..3;
is ($ws->{Cells}[0][ 4]{Format}{Font}{Color},    18, "E1 font color");

my %expect = (
    B7 => [ 1, 6, 39673, "ddd, dd mmm yyyy", "Mon, 13 Aug 2008" ],
    C7 => [ 2, 6, 39673, "m-d-yy",           "8-13-08"          ],
    );
# Test ddd, dd MM yyyy
foreach my $cell ("B7", "C7") {
    my ($c, $r) = @{$expect{$cell}}[0,1];
    ok (my $wc     = $ws->{Cells}[$r][$c],             "Cell   $cell");
    ok (my $fmt    = $wc->{Format},                    "Format $cell");
    ok (my $fmtstr = $bk->{FormatStr}{$fmt->{FmtIdx}}, "FmtStr $cell");
	   $fmtstr and $fmtstr =~ s{\\}{}g;	# Unescape for test
	   $fmtstr and $fmtstr =~ s{/}{-}g and
	       $expect{$cell}[4] =~ s{-}{/}g;	# System locale's deate-sep is used :(

    is ($wc->{Val}, $expect{$cell}[2], "$cell value");
    is ($fmtstr,    $expect{$cell}[3], "$cell Format string");
    is ($wc->Value, $expect{$cell}[4], "$cell formatted value");
    }
