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
    'Crypt::RC4'          => 0,
    'Digest::Perl::MD5'   => 0,
);

my %resources = (
    homepage  => 'http://github.com/runrig/spreadsheet-parseexcel/',
    repository  => 'http://github.com/runrig/spreadsheet-parseexcel',
    bugtracker  => 'http://github.com/runrig/spreadsheet-parseexcel/issues',
    MailingList => 'http://groups.google.com/group/spreadsheet-parseexcel',
);

my %params = (
    'AUTHOR'        => 'Douglas Wilson (dougw@cpan.org)',
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

if ( $ExtUtils::MakeMaker::VERSION ge '6.46' ) {
    $params{META_MERGE} = { resources => \%resources };
}

WriteMakefile(%params);
