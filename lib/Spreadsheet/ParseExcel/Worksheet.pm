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

our $VERSION = '0.48';

use Scalar::Util qw(weaken);

sub new {
    my ( $class, %rhIni ) = @_;
    my $self = \%rhIni;
    weaken $self->{_Book};

    $self->{Cells}       = undef;
    $self->{DefColWidth} = 8.38;
    bless $self, $class;
}

#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel::Worksheet->sheetNo
#------------------------------------------------------------------------------
sub sheetNo {
    my ($oSelf) = @_;
    return $oSelf->{_SheetNo};
}

#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel::Worksheet->Cell
#------------------------------------------------------------------------------
sub get_cell {
    my ( $oSelf, $iR, $iC ) = @_;

    # return undef if no arguments are given or if no cells are defined
    return
      if ( ( !defined($iR) )
        || ( !defined($iC) )
        || ( !defined( $oSelf->{MaxRow} ) )
        || ( !defined( $oSelf->{MaxCol} ) ) );

    # return undef if outside defined rectangle
    return
      if ( ( $iR < $oSelf->{MinRow} )
        || ( $iR > $oSelf->{MaxRow} )
        || ( $iC < $oSelf->{MinCol} )
        || ( $iC > $oSelf->{MaxCol} ) );

    # return the Cell object
    return $oSelf->{Cells}[$iR][$iC];
}
*Cell = *get_cell;

#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel::Worksheet->RowRange
#------------------------------------------------------------------------------
sub row_range {
    my ($oSelf) = @_;
    my $iMin = $oSelf->{MinRow} || 0;
    my $iMax = defined( $oSelf->{MaxRow} ) ? $oSelf->{MaxRow} : ( $iMin - 1 );

    # return the range
    return ( $iMin, $iMax );
}
*RowRange = *row_range;

#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel::Worksheet->ColRange
#------------------------------------------------------------------------------
sub col_range {
    my ($oSelf) = @_;
    my $iMin = $oSelf->{MinCol} || 0;
    my $iMax = defined( $oSelf->{MaxCol} ) ? $oSelf->{MaxCol} : ( $iMin - 1 );

    # return the range
    return ( $iMin, $iMax );
}
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
