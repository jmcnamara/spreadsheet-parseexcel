use strict;
use warnings;
use ExtUtils::MakeMaker;
use Config;

my %deps = (
    'OLE::Storage_Lite'   => 0.08,
    'IO::File'            => 0,
    'Scalar::Util'        => 0,
    'IO::Scalar'          => 0,
    'Test::More'          => '0.47',
);
if (exists $Config{useperlio} && $Config{useperlio} eq "define") {
    print "\$Config{useperlio} is defined, removing requirement of IO::Scalar\n";
#    delete $deps{'IO::Scalar'};
}

WriteMakefile(
          'PL_FILES'     => {},
          'INSTALLDIRS'  => 'site',
          'NAME'         => 'Spreadsheet::ParseExcel',
          'EXE_FILES'    => [],
          'VERSION_FROM' => 'lib/Spreadsheet/ParseExcel.pm',
          'PREREQ_PM'    => \%deps,

           test          => {TESTS => 't/*.t t/examples/*t'}

);
