package Spreadsheet::ParseExcel::Worksheet;
use strict;
use warnings;

our $VERSION = '0.33';

use overload 
    '0+'        => \&sheetNo,
    'fallback'  => 1,
;
use Scalar::Util qw(weaken);

sub new {
    my ($class, %rhIni) = @_;
    my $self = \%rhIni;
    weaken $self->{_Book};

    $self->{Cells}=undef;
    $self->{DefColWidth}=8.38;
    bless $self, $class;
}
#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel::Worksheet->sheetNo
#------------------------------------------------------------------------------
sub sheetNo {
    my($oSelf) = @_;
    return $oSelf->{_SheetNo};
}
#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel::Worksheet->Cell
#------------------------------------------------------------------------------
sub Cell {
    my($oSelf, $iR, $iC) = @_;

    # return undef if no arguments are given or if no cells are defined
    return  if ((!defined($iR)) || (!defined($iC)) ||
                (!defined($oSelf->{MaxRow})) || (!defined($oSelf->{MaxCol})));
    
    # return undef if outside defined rectangle
    return  if (($iR < $oSelf->{MinRow}) || ($iR > $oSelf->{MaxRow}) ||
                ($iC < $oSelf->{MinCol}) || ($iC > $oSelf->{MaxCol}));
    
    # return the Cell object
    return $oSelf->{Cells}[$iR][$iC];
}
#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel::Worksheet->RowRange
#------------------------------------------------------------------------------
sub RowRange {
    my($oSelf) = @_;
    my $iMin = $oSelf->{MinRow} || 0;
    my $iMax = defined($oSelf->{MaxRow}) ? $oSelf->{MaxRow} : ($iMin-1);

    # return the range
    return($iMin, $iMax);
}
#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel::Worksheet->ColRange
#------------------------------------------------------------------------------
sub ColRange {
    my($oSelf) = @_;
    my $iMin = $oSelf->{MinCol} || 0;
    my $iMax = defined($oSelf->{MaxCol}) ? $oSelf->{MaxCol} : ($iMin-1);

    # return the range
    return($iMin, $iMax);
}

#DESTROY {
#    my ($self) = @_;
#    warn "DESTROY $self called\n"
#}

1;
