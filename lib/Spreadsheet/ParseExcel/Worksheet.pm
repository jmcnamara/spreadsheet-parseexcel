package Spreadsheet::ParseExcel::Worksheet;

###############################################################################
#
# Spreadsheet::ParseExcel::Worksheet - A class for Worksheets.
#
# Used in conjunction with Spreadsheet::ParseExcel.
#
# Copyright (c) 2009      John McNamara
# Copyright (c) 2006-2008 Gabor Szabo
# Copyright (c) 2000-2006 Kawai Takanori
#
# perltidy with standard settings.
#
# Documentation after __END__
#

use strict;
use warnings;
use Scalar::Util qw(weaken);

our $VERSION = '0.48';

###############################################################################
#
# new()
#
sub new {

    my ( $class, %properties ) = @_;

    my $self = \%properties;

    weaken $self->{_Book};

    $self->{Cells}       = undef;
    $self->{DefColWidth} = 8.38;

    return bless $self, $class;
}

###############################################################################
#
# sheet_num()
#
sub sheet_num {

    my $self = shift;

    return $self->{_SheetNo};
}

###############################################################################
#
# get_cell()
#
sub get_cell {

    my ( $self, $row, $col ) = @_;

    if (   !defined $row
        || !defined $col
        || !defined $self->{MaxRow}
        || !defined $self->{MaxCol} )
    {

        # Return undef if no arguments are given or if no cells are defined.
        return undef;
    }
    elsif ($row < $self->{MinRow}
        || $row > $self->{MaxRow}
        || $col < $self->{MinCol}
        || $col > $self->{MaxCol} )
    {

        # Return undef if outside allowable row/col range.
        return undef;
    }
    else {

        # Return the Cell object.
        return $self->{Cells}->[$row]->[$col];
    }
}

###############################################################################
#
# row_range()
#
sub row_range {

    my $self = shift;

    my $min = $self->{MinRow} || 0;
    my $max = defined( $self->{MaxRow} ) ? $self->{MaxRow} : ( $min - 1 );

    return ( $min, $max );
}

###############################################################################
#
# col_range()
#
sub col_range {

    my $self = shift;

    my $min = $self->{MinCol} || 0;
    my $max = defined( $self->{MaxCol} ) ? $self->{MaxCol} : ( $min - 1 );

    return ( $min, $max );
}

###############################################################################
#
# Map legacy method names to new names.
#
*sheetNo  = *sheet_num;
*Cell     = *get_cell;
*RowRange = *row_range;
*ColRange = *col_range;

1;

__END__

=pod

=head1 NAME

Spreadsheet::ParseExcel::Worksheet - A class for Worksheets.

=head1 SYNOPSIS

See the documentation for Spreadsheet::ParseExcel.

=head1 DESCRIPTION

This module is used in conjunction with Spreadsheet::ParseExcel. See the documentation for Spreadsheet::ParseExcel.

=head1 AUTHOR

Maintainer 0.40+: John McNamara jmcnamara@cpan.org

Maintainer 0.27-0.33: Gabor Szabo szabgab@cpan.org

Original author: Kawai Takanori kwitknr@cpan.org

=head1 COPYRIGHT

Copyright (c) 2009 John McNamara

Copyright (c) 2006-2008 Gabor Szabo

Copyright (c) 2000-2006 Kawai Takanori

All rights reserved.

You may distribute under the terms of either the GNU General Public License or the Artistic License, as specified in the Perl README file.

=cut
