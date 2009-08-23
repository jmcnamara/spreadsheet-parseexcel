#!perl -w

use strict;
use Test::More tests => 14;

use utf8;
use Encode qw(encode);

use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtJapan;

my $xl   = Spreadsheet::ParseExcel->new();
my $fmtj = Spreadsheet::ParseExcel::FmtJapan->new(Code => 'sjis');

my $book = $xl->Parse("t/excel_files/TestEncoding.xls", $fmtj);
ok $book, "load TestEncoding.xls";

my $sheet = $book->worksheet(0);

my @expected = (
	['This is a test file for Japanese encoding'],
	[qw(ローマ数字	Ⅰ	Ⅱ)],
	[qw(丸数字	①	②)],
	[qw(年号	㍻	㍼)],
	[qw(その他	㈱	～)],
);


my($rmin, $rmax) = $sheet->row_range();
my($cmin, $cmax) = $sheet->col_range();

#binmode STDOUT, ':encoding(cp932)';

for my $i($rmin .. $rmax){
	for my $j($cmin .. $cmax){
		my $cell = $sheet->get_cell($i, $j);
		next unless $cell && $cell->value;
		#print $cell->value, "\n";
		is $cell->value, encode(cp932 => $expected[$i][$j]), "[$i, $j]";
	}
}
