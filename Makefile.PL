use strict;
use warnings;
use ExtUtils::MakeMaker;
use Config;
use 5.008;

my %deps = (
    'OLE::Storage_Lite'   => 0.19,
    'IO::File'            => 0,
    'Scalar::Util'        => 0,
    'IO::Scalar'          => 0,
);

if (exists $Config{useperlio} && $Config{useperlio} eq "define") {
#    delete $deps{'IO::Scalar'};
}

WriteMakefile(
    'AUTHOR'        => 'John McNamara (jmcnamara@cpan.org)',
    'NAME'          => 'Spreadsheet::ParseExcel',
    'ABSTRACT'      => 'Read information from an Excel file.',
    'VERSION_FROM'  => 'lib/Spreadsheet/ParseExcel.pm',
    'LICENSE'       => 'perl',
    'EXE_FILES'     => [],
    'INSTALLDIRS'   => 'site',
    'PL_FILES'      => {},
    'PREREQ_PM'     => \%deps,
    'NEEDS_LINKING' => 0,
    test            => { TESTS => 't/*.t' }
);

