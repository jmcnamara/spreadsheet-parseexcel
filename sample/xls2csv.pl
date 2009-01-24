#!/usr/bin/perl -w

# Script to extract a range of columns/rows from an Excel spreadsheet
# and present it as a csv, there is also the option to rotate the
# output, 
#
# (c) kevin Mulholland 2002, kevin@moodfarm.demon.co.uk
# This code is released under the Perl Artistic License

#
# See also the following utilities:.
#     Ken Prows' xls2csv:
#     http://search.cpan.org/~ken/xls2csv/script/xls2csv
#     H.Merijn Brand's xls2csv (which is part of Spreadsheet::Read):
#     http://search.cpan.org/~hmbrand/Spreadsheet-Read/
#

use strict ;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::Utility qw(xls2csv);
use Text::CSV_XS;

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
