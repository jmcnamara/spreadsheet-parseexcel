package Spreadsheet::ParseExcel::Workbook;

###############################################################################
#
# Spreadsheet::ParseExcel::Workbook - A class for Workbooks.
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

our $VERSION = '0.44';

sub new {
    my ($class) = @_;
    my $self = {};
    bless $self, $class;
}

#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel::Workbook->ParseAbort
#------------------------------------------------------------------------------
sub ParseAbort {
    my ( $self, $val ) = @_;
    $self->{_ParseAbort} = $val;
}

#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel::Workbook->Parse
#------------------------------------------------------------------------------
sub Parse {
    my ( $class, $source, $oFmt ) = @_;
    my $excel = Spreadsheet::ParseExcel->new;
    my $workbook = $excel->Parse( $source, $oFmt );
    $workbook->{_Excel} = $excel;
    return $workbook;
}

#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel::Workbook Worksheet
#------------------------------------------------------------------------------
sub Worksheet {
    my ( $oBook, $sName ) = @_;
    my $oWkS;
    foreach $oWkS ( @{ $oBook->{Worksheet} } ) {
        return $oWkS if ( $oWkS->{Name} eq $sName );
    }
    if ( $sName =~ /^\d+$/ ) {
        return $oBook->{Worksheet}->[$sName];
    }
    return undef;
}
*worksheet = *Worksheet;

#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel::Workbook worksheets
#------------------------------------------------------------------------------
sub worksheets {
    my $self = shift;

    return @{ $self->{Worksheet} };
}

1;

__END__

=pod

=head1 NAME

Spreadsheet::ParseExcel::Workbook - A class for Workbooks.

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
