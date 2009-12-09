#!/usr/bin/perl -w

use strict;
use warnings;
use ExtUtils::MakeMaker;
use 5.008;

my %deps = (
    'OLE::Storage_Lite'   => 0.19,
    'IO::File'            => 0,
    'Scalar::Util'        => 0,
    'IO::Scalar'          => 0,
);

my %resources = (
    repository  => 'http://github.com/jmcnamara/spreadsheet-parseexcel',
    MailingList => 'http://groups.google.com/group/spreadsheet-parseexcel',
);

my %params = (
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
    'test'          => { TESTS => 't/*.t' }
);

if ( $ExtUtils::MakeMaker::VERSION >= 6.46 ) {
    $params{META_MERGE} = { resources => \%resources };
}

WriteMakefile(%params);
