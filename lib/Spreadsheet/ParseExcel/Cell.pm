package Spreadsheet::ParseExcel::Cell;
use strict;
use warnings;

our $VERSION = '0.42';

sub new {
    my($sPkg, %rhKey)=@_;
    my($sWk, $iLen);
    my $self = \%rhKey;

    bless $self, $sPkg;
}

sub value {
    my($self)=@_;
    return $self->{_Value};
}

sub unformatted {
    my($self)=@_;
    return $self->{Val};
}

*Value = *value;

#DESTROY {
#    my ($self) = @_;
#    warn "DESTROY $self called\n"
#}

1;
