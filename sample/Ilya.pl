use strict;
use Spreadsheet::ParseExcel;
package IlyaPackage; #:-)
sub new($){
    my $self = shift;
    my $obj = {};
    bless $obj, $self;
}
sub cb_routine($@) {
    my($self, $oBook, $iSheet, $iRow, $iCol, $oCell) = @_;
    print "( $iRow , $iCol ) =>", $oCell->Value, "\n";
}
sub parse($$){
    my($self, $sFile) = @_;
    my $oEx = 
        Spreadsheet::ParseExcel->new(
            CellHandler => \&cb_routine, 
            Object => $self,
            NotSetCell => 1);
    $oEx->Parse($sFile);
}
1;
my $oIlya =IlyaPackage->new;
$oIlya->parse($ARGV[0]);
