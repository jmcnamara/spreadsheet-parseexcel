#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::ParseExcel.
#
# Tests for number format handling using FmtExcel(). See note below.
#
# reverse('©'), January 2009, John McNamara, jmcnamara@cpan.org
#

use strict;

use Spreadsheet::ParseExcel::Utility 'ExcelFmt';
use Test::More tests => 3;

###############################################################################
#
# Test cases for special cases or user supplied format issues.
#
my @testcases = (
    # No, Number,      Expected,       Format string,  TODO note (if any).
    [ 1,  0.01023148,  '12:14:44 AM',  'hh:mm:ss AM/PM' ],
    [ 1,  0.01024306,  '12:14:45 AM',  'hh:mm:ss AM/PM' ],
    [ 1,  0.01025463,  '12:14:46 AM',  'hh:mm:ss AM/PM' ],
);

###############################################################################
#
# Run tests.
#

for my $test_ref (@testcases) {

    my $number   = $test_ref->[1];
    my $expected = $test_ref->[2];
    my $format   = $test_ref->[3];
    my $got      = ExcelFmt( $format, $number );

    local $TODO  = $test_ref->[4] if defined $test_ref->[4];

    is( $got, $expected, " \tFormat = $format, Result = $got" );
}

__END__
