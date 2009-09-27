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

    if ( !defined $self->{Code} ) {
        return 1;
    }
    elsif ( $self->{Code} eq 'ucs2' ) {
        return 2;
    }
    elsif ( $self->{Code} eq '_native_' ) {
        return 3;
    }
    else {
        return 0;
    }

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

Formatted in this sense refers to the numeric fomat of the cell value. For example a number such as 40177 might be formatted as 40,117, 40117.000 or even as the date 2009/12/30.

If the cell doesn't contain a numeric format then the formatted and unformated cell values are the same, see the C<unformatted()> method below.

For a defined C<$cell> the C<value()> method will always return a value. In the case of a cell with formatting but no numeric or string contents the method will return the empyt string C<''>.


=head2 unformatted()

The C<unformatted()> method returns the unformatted value of the cell.

    my $unformatted = $cell->unformatted();

Returns the cell value without a numeric format. See the C<value()> method above.


=head2 get_format()

The C<get_format()> method returns the L<Spreadsheet::ParseExcel::Format> object for the cell.

    my $format = $cell->get_format();

If a user defined format hasn't been applied to the cell then the default cell format is returned.


=head2 type()

The C<type()> method returns the type of cell such as Text, Numeric or Date. If the type was detected as Numeric, and the Cell Format matches m{^[dmy][-\\/dmy]*$}i, it will be treated as a Date type.

    my $type = $cell->type();

See also L<Dates and Time in Excel>.


=head2 encoding()

The C<encoding()> method returns the character encoding of the cell. It is either undef, ucs2 or _native_.  If undef then the character encoding is generally ascii. If the cell encoding is C<_native_> it means that cell encoding is 'sjis' or something similar.

    my $encoding = $cell->encoding();

Returns 0 if the property isn't set.


=head2 is_merged()

The C<is_merged()> method returns true if the cell is merged.

    my $is_merged = $cell->is_merged();

Returns C<undef> if the property isn't set.


=head2 get_rich_text()

The C<get_rich_text()> method returns an array ref of font information about each string block in a "rich", i.e. multi-format, string. Each entry has the form: [ $start_position, $font_object ]

    my $rich_text = $cell->get_rich_text();

Returns 0 if the property isn't set.


=head1 Dates and Time in Excel

Dates and times in Excel are represented by real numbers, for example "Jan 1 2001 12:30 PM" is represented by the number 36892.521.

The integer part of the number stores the number of days since the epoch and the fractional part stores the percentage of the day.

A date or time in Excel is just like any other number. To have the number display as a date you must apply an Excel number format to it. Here are some examples.

    Number     Format                      Result
    39506.5    'dd/mm/yy'                  28/02/08
    39506.5    'mm/dd/yy'                  02/28/08
    39506.5    'd-m-yyyy'                  28-2-2008
    39506.5    'dd/mm/yy hh:mm'            28/02/08 12:00
    39506.5    'd mmm yyyy'                28 Feb 2008
    39506.5    'mmm d yyyy hh:mm AM/PM'    Feb 28 2008 12:00 PM


TODO ParseExcel date handling functions.

For date conversions using the CPAN C<DateTime> framework see L<DateTime::Format::Excel> http://search.cpan.org/search?dist=DateTime-Format-Excel


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
