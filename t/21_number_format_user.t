#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::ParseExcel.
#
# Tests for number format handling using FmtExcel(). See note below.
#
# reverse('�'), January 2009, John McNamara, jmcnamara@cpan.org
#

use strict;

use Spreadsheet::ParseExcel::Utility 'ExcelFmt';
use Test::More tests => 27;

###############################################################################
#
# Test cases for special cases or user supplied format issues.
#
my @testcases = (

    # No, Number,      Expected,       Format string,  TODO note (if any).

    # Test for invalid 12-hour clock values.
    # http://rt.cpan.org/Public/Bug/Display.html?id=41192
    [ 1,  0.01023148,  '12:14:44 AM',  'hh:mm:ss AM/PM' ],
    [ 2, 0.01024306, '12:14:45 AM', 'hh:mm:ss AM/PM' ],
    [ 3, 0.01025463, '12:14:46 AM', 'hh:mm:ss AM/PM' ],

    # Tests for upper case formats from OpenOffice.org.
    # http://rt.cpan.org/Public/Bug/Display.html?id=20526
    # http://rt.cpan.org/Public/Bug/Display.html?id=31206
    # http://rt.cpan.org/Public/Bug/Display.html?id=40307
    [ 4, 37653.521,  '2/1/03',      'M/D/YY' ],
    [ 5, 37653.521,  '02/01/2003',  'MM/DD/YYYY' ],
    [ 6, 37653.521,  '01/02/2003',  'DD/MM/YYYY' ],
    [ 7, 37653.521,  '20030201',    'YYYYMMDD' ],
    [ 8, 37653.521,  '2003-02-01',  'YYYY-MM-DD' ],
    [ 9, 0.01023148, '12:14:44 AM', 'HH:MM:SS AM/PM' ],

    # Tests for overflow hours and minutes formats.
    [ 10, 0.4, '9:36:00',  '[h]:mm:ss' ],
    [ 11, 1.4, '33:36:00', '[h]:mm:ss' ],
    [ 12, 2.4, '57:36:00', '[h]:mm:ss' ],
    [ 13, 0.6, '14:24:00', '[h]:mm:ss' ],
    [ 14, 1.6, '38:24:00', '[h]:mm:ss' ],
    [ 15, 2.6, '62:24:00', '[h]:mm:ss' ],
    [ 16, 0.4, 9,          '[h]' ],
    [ 17, 1.4, 33,         '[h]' ],
    [ 18, 2.4, 57,         '[h]' ],
    [ 19, 0.4, 576,        '[mm]' ],
    [ 20, 1.4, 2016,       '[mm]' ],
    [ 21, 2.4, 3456,       '[mm]' ],

    # Formats that don't overflow. Counter examples of the above.
    [ 22, 0.4, '9:36:00',  'h:mm:ss' ],
    [ 23, 1.4, '9:36:00',  'h:mm:ss' ],
    [ 24, 2.4, '9:36:00',  'h:mm:ss' ],
    [ 25, 0.6, '14:24:00', 'h:mm:ss' ],
    [ 26, 1.6, '14:24:00', 'h:mm:ss' ],
    [ 27, 2.6, '14:24:00', 'h:mm:ss' ],
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

    is( $got, $expected, " \tFormat = $format,\tResult = $got" );
}

__END__
