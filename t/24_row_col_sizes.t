#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::ParseExcel.
#
# Tests for the row and column size conversions.
#
# reverse('©'), August 2009, John McNamara, jmcnamara@cpan.org
#

use strict;

use Spreadsheet::ParseExcel;
use Test::More tests => 37;

###############################################################################
#
# Test cases for row/column sizes extracted from an Excel workbook.
#
my @col_testcases = (
    # Pix      Internal   User
    [ 1,       36,        0.08   ],
    [ 2,       73,        0.17   ],
    [ 3,       109,       0.25   ],
    [ 4,       146,       0.33   ],
    [ 5,       182,       0.42   ],
    [ 6,       219,       0.50   ],
    [ 7,       256,       0.58   ],
    [ 8,       292,       0.67   ],
    [ 9,       329,       0.75   ],
    [ 10,      365,       0.83   ],
    [ 11,      402,       0.92   ],
    [ 12,      438,       1.00   ],
    [ 13,      475,       1.14   ],
    [ 14,      512,       1.29   ],
    [ 15,      548,       1.43   ],
    [ 16,      585,       1.57   ],
    [ 17,      621,       1.71   ],
    [ 18,      658,       1.86   ],
    [ 19,      694,       2.00   ],
    [ 20,      731,       2.14   ],
    [ 21,      768,       2.29   ],
    [ 22,      804,       2.43   ],
    [ 23,      841,       2.57   ],
    [ 24,      877,       2.71   ],
    [ 25,      914,       2.86   ],
    [ 26,      950,       3.00   ],
    [ 27,      987,       3.14   ],
    [ 28,      1024,      3.29   ],
    [ 29,      1060,      3.43   ],
    [ 30,      1097,      3.57   ],
    [ 64,      2340,      8.43   ],
    [ 399,     14592,     56.29  ],
    [ 400,     14628,     56.43  ],
    [ 401,     14665,     56.57  ],
    [ 999,     36534,     142.00 ],
    [ 1000,    36571,     142.14 ],
    [ 1001,    36608,     142.29 ],
);

# This test data isn't used for now since the row conversion is straightforward.
my @row_testcases = (
    # Pix      Internal   User
    [ 1,       15,        0.75  ],
    [ 3,       45,        2.25  ],
    [ 4,       60,        3     ],
    [ 5,       75,        3.75  ],
    [ 7,       105,       5.25  ],
    [ 8,       120,       6     ],
    [ 9,       135,       6.75  ],
    [ 15,      225,       11.25 ],
    [ 16,      240,       12    ],
    [ 17,      255,       12.75 ],
    [ 31,      465,       23.25 ],
    [ 32,      480,       24    ],
    [ 33,      495,       24.75 ],
    [ 63,      945,       47.25 ],
    [ 64,      960,       48    ],
    [ 65,      975,       48.75 ],
);


###############################################################################
#
# Run tests.
#
my $parser = Spreadsheet::ParseExcel->new();

for my $test_ref (@col_testcases) {

    my $excel_width = $test_ref->[1];
    my $user_width  = $test_ref->[2];
    my $got         = $parser->_convert_col_width($excel_width);

    is( $got, $user_width );
}

__END__
