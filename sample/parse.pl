#!/usr/bin/perl
use strict;
use warnings;

use Spreadsheet::ParseExcel;
use Getopt::Long qw(GetOptions);

my $file;
my $help;
my $dump;

my ($row, $col);

GetOptions(
    "file=s" => \$file,
    "help"   => \$help,

    "dump"   => \$dump,
    "row=i"  => \$row,
    "col=i"  => \$col,
) or usage();
usage() if $help;
usage() if not $file;

my $excel = Spreadsheet::ParseExcel::Workbook->Parse($file);
if ($dump) {
    foreach my $sheet (@{$excel->{Worksheet}}) {
        printf("Sheet: %s\n", $sheet->{Name});
        $sheet->{MaxRow} ||= $sheet->{MinRow};
        foreach my $row ($sheet->{MinRow} .. $sheet->{MaxRow}) {
            $sheet->{MaxCol} ||= $sheet->{MinCol};
            foreach my $col ($sheet->{MinCol} ..  $sheet->{MaxCol}) {
                my $cell = $sheet->{Cells}[$row][$col];
                if ($cell) {
                    printf("( %s , %s ) => %s\n", $row, $col, $cell->{Val});
                }
            }
        }
    }
}
if (defined $row and defined $col) {
    foreach my $sheet (@{$excel->{Worksheet}}) {
        printf("Sheet: %s\n", $sheet->{Name});
        my $cell = $sheet->{Cells}[$row][$col];
        printf("( %s , %s ) => '%s'\n", $row, $col, $cell->{Val});
        printf("( %s , %s ) => '%s'\n", $row, $col, $cell->Value);
    }
}


sub usage {
    print <<"END_USAGE";
Usage: $0
        --file FILENAME
        --dump

        --row  ROW
        --col  COL

        --help
END_USAGE
    exit;
}

