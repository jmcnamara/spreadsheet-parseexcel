package Spreadsheet::ParseExcel::Workbook;
use strict;
use warnings;

our $VERSION = '0.42';

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

#------------------------------------------------------------------------------
# Spreadsheet::ParseExcel::Workbook worksheets
#------------------------------------------------------------------------------
sub worksheets {
    my $self = shift;

    return @{ $self->{Worksheet} };
}

#DESTROY {
#    my ($self) = @_;
#    warn "DESTROY $self called\n"
#}

1;
