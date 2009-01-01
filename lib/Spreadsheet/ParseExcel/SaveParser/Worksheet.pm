package Spreadsheet::ParseExcel::SaveParser::Worksheet;
use strict;
use warnings;

#==============================================================================
# Spreadsheet::ParseExcel::SaveParser::Worksheet
#==============================================================================

use base 'Spreadsheet::ParseExcel::Worksheet';
our $VERSION = '0.33';


sub new {
  my ($sClass, %rhIni) = @_;
  $sClass->SUPER::new(%rhIni);  # returns object
}
#------------------------------------------------------------------------------
# AddCell (for Spreadsheet::ParseExcel::SaveParser::Worksheet)
#------------------------------------------------------------------------------
sub AddCell {
    my($oSelf, $iR, $iC, $sVal, $oCell, $sCode)=@_;
    $oSelf->{_Book}->AddCell($oSelf->{_SheetNo}, $iR, $iC, $sVal, $oCell, $sCode);
}
#------------------------------------------------------------------------------
# Protect (for Spreadsheet::ParseExcel::SaveParser::Worksheet)
#  - Password = undef   ->  No protect
#  - Password = ''      ->  Protected. No password
#  - Password = $pwd    ->  Protected. Password = $pwd
#------------------------------------------------------------------------------
sub Protect {
    my($oSelf, $sPassword)=@_;
    $oSelf->{Protect} = $sPassword;
}

1;
