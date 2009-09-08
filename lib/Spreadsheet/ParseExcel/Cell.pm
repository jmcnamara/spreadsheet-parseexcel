package Spreadsheet::ParseExcel::Cell;

###############################################################################
#
# Spreadsheet::ParseExcel::Cell - A class for Cell data and formatting.
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

###############################################################################
#
# new()
#
# Constructor.
#
sub new {
    my ( $package, %properties ) = @_;
    my $self = \%properties;

    bless $self, $package;
}

###############################################################################
#
# value()
#
# Returns the formatted value of the cell.
#
sub value {

    my $self = shift;

    return $self->{_Value};
}

###############################################################################
#
# unformatted()
#
# Returns the unformatted value of the cell.
#
sub unformatted {

    my $self = shift;

    return $self->{Val};
}

###############################################################################
#
# get_format()
#
# Returns the Format object for the cell.
#
sub get_format {

    my $self = shift;

    return $self->{Format};
}

###############################################################################
#
# type()
#
# Returns the type of cell such as Text, Numeric or Date.
#
sub type {

    my $self = shift;

    return $self->{Type};
}

###############################################################################
#
# encoding()
#
# Returns the character encoding of the cell.
#
sub encoding {

    my $self = shift;

    return $self->{Code};
}

###############################################################################
#
# is_merged()
#
# Returns true if the cell is merged.
#
sub is_merged {

    my $self = shift;

    return $self->{Merged};
}

###############################################################################
#
# get_rich_text()
#
# Returns an array ref of font information about each string block in a "rich",
# i.e. multi-format, string.
#
sub get_rich_text {

    my $self = shift;

    return $self->{Rich};
}

###############################################################################
#
# Mapping between legacy method names and new names.
#
{
    no warnings;    # Ignore warnings about variables used only once.
    *Value = *value;
}

1;

__END__

=pod

=head1 NAME

Spreadsheet::ParseExcel::Cell - A class for Cell data and formatting.

=head1 SYNOPSIS

See the documentation for Spreadsheet::ParseExcel.

=head1 DESCRIPTION

This module is used in conjunction with Spreadsheet::ParseExcel. See the documentation for Spreadsheet::ParseExcel.

=head1 Methods

The following Cell methods are available:

    $cell->value()
    $cell->unformatted()
    $cell->get_format()
    $cell->type()
    $cell->encoding()
    $cell->is_merged()
    $cell->get_rich_text()


=head2 value()

The C<value()> method returns the formatted value of the cell.

    my $value = $cell->value();

Returns 0 if the property isn't set.


=head2 unformatted()

The C<unformatted()> method returns the unformatted value of the cell.

    my $unformatted = $cell->unformatted();

Returns 0 if the property isn't set.


=head2 get_format()

The C<get_format()> method returns the L<"Format"> object for the cell.

    my $format = $cell->get_format();

Returns 0 if the property isn't set.


=head2 type()

The C<type()> method returns the type of cell such as Text, Numeric or Date. If the type was detected as Numeric, and the Cell Format matches m{^[dmy][-\\/dmy]*$}, it will be treated as a Date type.

    my $type = $cell->type();

Returns 0 if the property isn't set.


=head2 encoding()

The C<encoding()> method returns the character encoding of the cell. It is either undef, ucs2 or _native_.  If undef then the character encoding is generally ascii. If the cell encoding is C<_native_> it means that cell encoding is 'sjis' or something similar.

    my $encoding = $cell->encoding();

Returns 0 if the property isn't set.


=head2 is_merged()

The C<is_merged()> method returns true if the cell is merged.

    my $is_merged = $cell->is_merged();

Returns 0 if the property isn't set.


=head2 get_rich_text()

The C<get_rich_text()> method returns an array ref of font information about each string block in a "rich", i.e. multi-format, string. Each entry has the form: [ $start_position, $font_object ]

    my $rich_text = $cell->get_rich_text();

Returns 0 if the property isn't set.








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
