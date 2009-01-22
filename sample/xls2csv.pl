#!/usr/bin/perl -w
# script to extract a range of columns/rows from an excel spreadsheet
# and present it as a csv, there is also the option to rotate the
# output, 
#
# (c) kevin Mulholland 2002, kevin@moodfarm.demon.co.uk
# this code is released under the Perl Artistic License

#
# Note: this utility doesn't currently handle embedded commas in extracted data.
# Use Ken Prows' xls2cvs http://search.cpan.org/~ken/xls2csv-1.06/script/xls2csv
# or H.Merijn Brand's xls2cvs (which is part of Spreadsheet::Read
# http://search.cpan.org/~hmbrand/Spreadsheet-Read/) instead.
# Both of these use Text::CSV_XS.
#

use strict ;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::Utility qw(xls2csv);

if(!(defined $ARGV[0])) {
	usage( "Bad Args") ;
	exit;
}

my $rotate = 0 ;
my $filename = $ARGV[0] ;
my $coords = $ARGV[1] ;

$rotate = 1 if( defined $ARGV[2] && $ARGV[2] eq "-rotate") ;

if( !$coords) {
   usage( "No co-ordinates defined") ;
   exit ;
}

if( ! -f $filename) {
   usage( "File $filename does not exist") ;
   exit ;
}

printf xls2csv( $filename, $coords, $rotate) ;

# -----------------------------------------------------------------------------
### error
# writes a message to STDERR
#
sub error {
   printf STDERR shift ;
}
# -----------------------------------------------------------------------------
### usage
# this decribes how the program as a whole is to be used
#
sub usage {
	my $errmsg = shift ;

	error( "\n" . $errmsg . "\n") if( $errmsg) ;

	error( "
Usage: $0 filename sheet-colrow:colrow [-rotate]
   eg: $0 filename.xls 1-A1:B12
       $0 filename.xls A1:M1 -rotate\n\n") ;
}
