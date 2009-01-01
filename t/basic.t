#!/usr/bin/perl
use strict;
use warnings;

use Test::More tests => 8;

use_ok('Spreadsheet::ParseExcel');
use_ok('Spreadsheet::ParseExcel::Dump');
use_ok('Spreadsheet::ParseExcel::FmtDefault');
use_ok('Spreadsheet::ParseExcel::Utility');
foreach my $mod (
            'OLE::Storage_Lite',
            'IO::File',
            'Scalar::Util',
            'IO::Scalar',
            ) {
   (my $path = $mod) =~ s{::}{/}g;
   my $str = sprintf("%-20s  ", $mod);
   if ($INC{"$path.pm"}) {
       $str .= sprintf("%s %s", $mod->VERSION, $INC{"$path.pm"});
   } else {
       $str .= "     not in memory";
   }
   diag $str;
}


eval "use  Jcode";
my $no_jcode = $@;

eval "use Unicode::Map";
my $no_unicode_map = $@;

eval "use Spreadsheet::WriteExcel";
my $no_writeexcel = $@;

SKIP: {
    skip "Need Jcode for additional tests", 1 if $no_jcode;
    use_ok('Spreadsheet::ParseExcel::FmtJapan');
}

SKIP: {
    skip "Need Unicode::Map for additional tests", 1 if $no_unicode_map;
    use_ok('Spreadsheet::ParseExcel::FmtUnicode');
}

SKIP: {
    skip "Need Jcode and Unicode::Map for additional tests", 1 if $no_jcode or $no_unicode_map;
    use_ok('Spreadsheet::ParseExcel::FmtJapan2');
}

SKIP: {
    skip "Need Spreadsheet::WriteExcel for additional tests", 1 if $no_writeexcel;
    use_ok('Spreadsheet::ParseExcel::SaveParser');
}


