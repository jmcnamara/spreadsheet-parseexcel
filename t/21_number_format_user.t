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

use Spreadsheet::ParseExcel::Utility qw( ExcelFmt LocaltimeExcel );
use Test::More tests => 146;

my $is_1904 = 1;

###############################################################################
#
# Test cases for special cases or user supplied format issues.
#
my @testcases = (

    # No, Number, Expected, Format string, Todo

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

    # Test for the examples in the Utility.pm docs.
    [ 28, 1234.567,  '$1,234.57',           '$#,##0.00' ],
    [ 29, 36892.521, '1 Jan 2001 12:30 PM', 'd mmm yyyy h:mm AM/PM' ],
    [ 30, LocaltimeExcel( 0, 0, 0, 1, 0, 101 ), '1 Jan 2001', 'd mmm yyyy' ],

    # Tests for locale in format string.
    # http://rt.cpan.org/Public/Bug/Display.html?id=43638
    [ 31, 39814, '1/1/09', 'm/d/yy;@' ],
    [ 32, 39845, '2/1/09', 'm/d/yy;@' ],
    [ 33, 39814, 'Jan-09', '[$-409]mmm-yy' ],
    [ 34, 39845, 'Feb-09', '[$-409]mmm-yy' ],

    # Tests for three part format strings.
    # http://rt.cpan.org/Public/Bug/Display.html?id=45009
    [ 35, 5,  '5.00', '0.00;(0.0);0%' ],
    [ 36, 0,  '0%',   '0.00;(0.0);0%' ],
    [ 37, -3, '-3.0', '0.00;(0.0);0%' ],

    # Tests for ignoring of all dots except the first in format strings.
    # http://rt.cpan.org/Public/Bug/Display.html?id=45502
    [ 38, 3.5008, '3.5008 oz.', '#.####\ \o\z.' ],
    [ 39, 3.5008, '3.5.0.0.8',  '#.#.#.#.#' ],

    # Tests for rounding.
    # http://rt.cpan.org/Public/Bug/Display.html?id=45626
    [ 40, 0.05,   '0.1',    '0.0' ],
    [ 41, 0.15,   '0.2',    '0.0' ],
    [ 42, 0.25,   '0.3',    '0.0' ],
    [ 43, 0.35,   '0.4',    '0.0' ],
    [ 44, 0.45,   '0.5',    '0.0' ],
    [ 45, 0.55,   '0.6',    '0.0' ],
    [ 46, 0.65,   '0.7',    '0.0' ],
    [ 47, 0.75,   '0.8',    '0.0' ],
    [ 48, 0.85,   '0.9',    '0.0' ],
    [ 49, 0.95,   '1.0',    '0.0' ],
    [ 50, 0.005,  '0.01',   '0.00' ],
    [ 51, 0.015,  '0.02',   '0.00' ],
    [ 52, 0.025,  '0.03',   '0.00' ],
    [ 53, 0.035,  '0.04',   '0.00' ],
    [ 54, 0.045,  '0.05',   '0.00' ],
    [ 55, 0.055,  '0.06',   '0.00' ],
    [ 56, 0.065,  '0.07',   '0.00' ],
    [ 57, 0.075,  '0.08',   '0.00' ],
    [ 58, 0.085,  '0.09',   '0.00' ],
    [ 59, 0.095,, '0.10',   '0.00' ],
    [ 60, 0.0005, '0.001',  '0.000' ],
    [ 61, 0.0015, '0.002',  '0.000' ],
    [ 62, 0.0025, '0.003',  '0.000' ],
    [ 63, 0.0035, '0.004',  '0.000' ],
    [ 64, 0.0045, '0.005',  '0.000' ],
    [ 65, 0.0055, '0.006',  '0.000' ],
    [ 66, 0.0065, '0.007',  '0.000' ],
    [ 67, 0.0075, '0.008',  '0.000' ],
    [ 68, 0.0085, '0.009',  '0.000' ],
    [ 69, 0.0095, '0.010',  '0.000' ],
    [ 70, 0.0005, '0.0005', '0.0000' ],
    [ 71, 0.0015, '0.0015', '0.0000' ],
    [ 72, 0.0025, '0.0025', '0.0000' ],
    [ 73, 0.0035, '0.0035', '0.0000' ],
    [ 74, 0.0045, '0.0045', '0.0000' ],
    [ 75, 0.0055, '0.0055', '0.0000' ],
    [ 76, 0.0065, '0.0065', '0.0000' ],
    [ 77, 0.0075, '0.0075', '0.0000' ],
    [ 78, 0.0085, '0.0085', '0.0000' ],
    [ 79, 0.0095, '0.0095', '0.0000' ],

    # Tests for valid dates.
    # http://rt.cpan.org/Public/Bug/Display.html?id=48831
    [ 80, 2958465, '31/12/9999', 'dd/mm/yyyy' ],
    [ 81, 2958466, '2958466', 'dd/mm/yyyy' ],
    [ 82, 4030433048023, '4030433048023', 'dd/mm/yyyy' ],
    [ 83, -1, '-1', 'dd/mm/yyyy' ],
    [ 84, 2957003, '31/12/9999', 'dd/mm/yyyy', undef, $is_1904 ],
    [ 85, 2957004, '2957004', 'dd/mm/yyyy',  undef, $is_1904],


    # Tests for real names for days and months.
    [ 86,  36528, 'Mon',    'ddd' ],
    [ 87,  36529, 'Tue',   'ddd' ],
    [ 88,  36530, 'Wed', 'ddd' ],
    [ 89,  36531, 'Thu',  'ddd' ],
    [ 90,  36532, 'Fri',    'ddd' ],
    [ 91,  36533, 'Sat',  'ddd' ],
    [ 92,  36534, 'Sun',    'ddd' ],
    [ 93,  36535, 'Monday',    'dddd' ],
    [ 94,  36536, 'Tuesday',   'dddd' ],
    [ 95,  36537, 'Wednesday', 'dddd' ],
    [ 96,  36538, 'Thursday',  'dddd' ],
    [ 97,  36539, 'Friday',    'dddd' ],
    [ 98,  36540, 'Saturday',  'dddd' ],
    [ 99,  36541, 'Sunday',    'dddd' ],
    [ 100, 36526, 'Jan',       'mmm' ],
    [ 101, 36557, 'Feb',       'mmm' ],
    [ 102, 36586, 'Mar',       'mmm' ],
    [ 103, 36617, 'Apr',       'mmm' ],
    [ 104, 36647, 'May',       'mmm' ],
    [ 105, 36678, 'Jun',       'mmm' ],
    [ 106, 36708, 'Jul',       'mmm' ],
    [ 107, 36739, 'Aug',       'mmm' ],
    [ 108, 36770, 'Sep',       'mmm' ],
    [ 109, 36800, 'Oct',       'mmm' ],
    [ 110, 36831, 'Nov',       'mmm' ],
    [ 111, 36861, 'Dec',       'mmm' ],
    [ 112, 36526, 'January',   'mmmm' ],
    [ 113, 36557, 'February',  'mmmm' ],
    [ 114, 36586, 'March',     'mmmm' ],
    [ 115, 36617, 'April',     'mmmm' ],
    [ 116, 36647, 'May',       'mmmm' ],
    [ 117, 36678, 'June',      'mmmm' ],
    [ 118, 36708, 'July',      'mmmm' ],
    [ 119, 36739, 'August',    'mmmm' ],
    [ 120, 36770, 'September', 'mmmm' ],
    [ 121, 36800, 'October',   'mmmm' ],
    [ 122, 36831, 'November',  'mmmm' ],
    [ 123, 36861, 'December',  'mmmm' ],
    [ 124, 36526, 'J',         'mmmmm' ],
    [ 125, 36557, 'F',         'mmmmm' ],
    [ 126, 36586, 'M',         'mmmmm' ],
    [ 127, 36617, 'A',         'mmmmm' ],
    [ 128, 36647, 'M',         'mmmmm' ],
    [ 129, 36678, 'J',         'mmmmm' ],
    [ 130, 36708, 'J',         'mmmmm' ],
    [ 131, 36739, 'A',         'mmmmm' ],
    [ 132, 36770, 'S',         'mmmmm' ],
    [ 133, 36800, 'O',         'mmmmm' ],
    [ 134, 36831, 'N',         'mmmmm' ],
    [ 135, 36861, 'D',         'mmmmm' ],
    [ 136, 1,     'Sun',       'ddd' ],
    [ 137, 127,   'Sun',       'ddd' ],
    [ 138, 36898, 'Sun',       'ddd' ],
    [ 139, 2958103,'Sun',       'ddd' ],


    # http://rt.cpan.org/Public/Bug/Display.html?id=60547
    [ 140, 27400,     '$27,400.00',      '[$$-409]#,##0.00'      ],
    [ 141, 826331.94, '826,331.94 руб.', '#,##0.00\ [$руб.-419]' ],
    [ 142, 826331.94, '826,331.94 RUR',  '#,##0.00\ [$RUR]'      ],

    # http://rt.cpan.org/Public/Bug/Display.html?id=93142
    [ 143, 41700.18, 'Sunday, March 02, 2014', 'dddd, mmmm dd, yyyy' ],
    [ 144, 41700.18, 'Sunday, March 02, 2014', '[$-F800]dddd, mmmm dd, yyyy' ],

    # http://rt.cpan.org/Ticket/Display.html?id=47072
    [ 145, 39814, '1/1/2009 12:00 AM', 'm/d/yyyy h:mm AM/PM' ],
    [ 146, '39813.999999994212963', '1/1/2009 12:00 AM', 'm/d/yyyy h:mm AM/PM' ],

);

###############################################################################
#
# Run tests.
#

for my $test_ref (@testcases) {

    my $number   = $test_ref->[1];
    my $expected = $test_ref->[2];
    my $format   = $test_ref->[3];
    my $is_1904  = $test_ref->[5];
    my $got      = ExcelFmt( $format, $number, $is_1904 );

    local $TODO  = $test_ref->[4] if defined $test_ref->[4];

    is( $got, $expected, " \t$number\t+ '$format'\t= $got" );
}

__END__
