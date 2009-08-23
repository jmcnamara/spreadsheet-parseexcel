#!perl -w

use strict;
use Test::More tests => 22;

use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtJapan;

use utf8;

my $xl   = Spreadsheet::ParseExcel->new();
my $fmtj = Spreadsheet::ParseExcel::FmtJapan->new();


my $book = $xl->Parse("t/excel_files/Test2000J.xls", $fmtj);
ok $book, "load Test2000J-Nengo.xls";

my $sheet = $book->worksheet(0);

my @expected = (
	['This is a test file for Japanese format', '', ''],
	[qw(明治	明治33年11月21日	M33.12.21)],
	[qw(大正	大正9年11月22日	T3.12.22)],
	[qw(昭和	昭和5年11月23日	S5.12.23)],
	[qw(平成	平成12年11月24日	H12.12.24)],
	[qw(日付	2009年7月1日 7月1日)],
	[qw(時刻	12時23分45秒 12時23分)],
);


my($rmin, $rmax) = $sheet->row_range();
my($cmin, $cmax) = $sheet->col_range();

#binmode STDOUT, ':encoding(cp932)';

for my $i($rmin .. $rmax){
	for my $j($cmin .. $cmax){
		#print $sheet->get_cell($i, $j)->value, "\n";
		is $sheet->get_cell($i, $j)->value, $expected[$i][$j], "[$i, $j]";
	}
}
