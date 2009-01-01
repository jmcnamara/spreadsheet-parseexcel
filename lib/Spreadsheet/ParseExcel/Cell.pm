package Spreadsheet::ParseExcel::Cell;
use strict;
use warnings;

our $VERSION = '0.33';

sub new {
    my($sPkg, %rhKey)=@_;
    my($sWk, $iLen);
    my $self = \%rhKey;

    bless $self, $sPkg;
}

sub Value {
    my($self)=@_;
    return $self->{_Value};
}
#DESTROY {
#    my ($self) = @_;
#    warn "DESTROY $self called\n"
#}

1;
