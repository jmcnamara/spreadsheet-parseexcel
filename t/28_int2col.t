#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::ParseExcel.
#
# Tests for Utility int2col() and col2int() functions..
#
# reverse('©'), August 2009, John McNamara, jmcnamara@cpan.org
#

use strict;

use Spreadsheet::ParseExcel::Utility qw( int2col col2int );
use Test::More tests => 3;

##############################################################################
#
# Tests.
#
my $col = 'A';
my @got_col;
my @got_int;
my @expected_col;
my @expected_int;

for my $int ( 0 .. 255 ) {
    $expected_col[$int] = $col;
    $expected_int[$int] = $int;

    $got_col[$int] = int2col($int);
    $got_int[$int] = col2int($col);

    $col++;
}

# General tests for full column range.
is_deeply( \@got_col, \@expected_col );
is_deeply( \@got_int, \@expected_int );

# Test for int2col in list context. RT 48967
my ($got) = int2col(27);
my $expected = 'AB';

is($got, $expected);

__END__
