#!perl -w

use strict;
use Test::More tests => 66;

use utf8;

use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtJapan;

my $xl   = Spreadsheet::ParseExcel->new();
my $fmtj = Spreadsheet::ParseExcel::FmtJapan->new();

foreach my $xls(qw(Test95J.xls Test97J.xls)){

	my $book = $xl->Parse("t/excel_files/$xls", $fmtj);
	ok $book, "load $xls";

	my $sheet = $book->worksheet(0);

	is $sheet->{Name}, 'Sheet1-ASC', '1. ASCII name';

	my @expected = (
		[ASC     => q{This Data is 'ASC Only'}],
		[Date    => q{1964/3/23}],
		[INTEGER => 12345],
		[Float   => 1.29],
		[Double  => 1234567.89012345],
		[Formula => 1246912.89012345],
		[Data    => 1234567.89],

		['BIG INTEGER'  => 123456789012],
	);

	#binmode STDOUT, ':encoding(cp932)';

	my($rmin, $rmax) = $sheet->row_range();
	my($cmin, $cmax) = $sheet->col_range();

	for my $i($rmin .. $rmax){
		for my $j($cmin .. $cmax){
			#print $sheet->get_cell($i, $j)->value, "\n";
			my $cell     = $sheet->get_cell($i, $j);
			my $got		 = $cell->value;
			my $expected = $expected[$i][$j];
			my $caption	 = "[$i, $j]";

			if ($expected =~ /\d\.\d+/) {
				_is_float($got, $expected, $caption);
			}
			else {
				is $got, $expected, $caption;
			}
		}
	}

	$sheet = $book->worksheet(1);

	is $sheet->{Name}, '漢字名', '2. Kanji name';

	@expected = (
		[ASC     => q{This Data is 'ASC Only'}],
		['漢字も入る' => '漢字のデータ'],
		[Date    => q{1964/3/23}],
		[INTEGER => 12345],
		[Float   => 1.29],
		[Double  => 1234567.89012345],
		[Formula => 1246912.89012345],
		[Float   => undef],
	);

	($rmin, $rmax) = $sheet->row_range();
	($cmin, $cmax) = $sheet->col_range();

	for my $i($rmin .. $rmax){
		for my $j($cmin .. $cmax){
			#print $sheet->get_cell($i, $j)->value, "\n";
			my $cell     = $sheet->get_cell($i, $j);
			my $got		 = ref($cell) ? $cell->value : $cell;
			my $expected = $expected[$i][$j];
			my $caption	 = "[$i, $j]";

			if (defined $expected && $expected =~ /\d\.\d+/) {
				_is_float($got, $expected, $caption);
			}
			else {
				is $got, $expected, $caption;
			}
		}
	}
}


###############################################################################
#
# _is_float()
#
# Helper function for float comparison. This is mainly to prevent failing tests
# on 64bit systems with extended doubles where the 128bit precision is compared
# against Excel's 64bit precision.
#
sub _is_float {

	my ( $got, $expected, $caption ) = @_;

	my $max = 1;
	$max = abs($got)	  if abs($got) > $max;
	$max = abs($expected) if abs($expected) > $max;

	if ( abs( $got - $expected ) <= 1e-15 * $max ) {
		ok( 1, $caption );
	}
	else {
		is( $got, $expected, $caption );
	}
}

