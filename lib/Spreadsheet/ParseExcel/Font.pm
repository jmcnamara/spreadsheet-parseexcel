package Spreadsheet::ParseExcel::Font;
use strict;
use warnings;

our $VERSION = '0.33';

sub new {
    my($class, %rhIni) = @_;
    my $self = \%rhIni;

    bless $self, $class;
}

#DESTROY {
#    my ($self) = @_;
#    warn "DESTROY $self called\n"
#}


1;
