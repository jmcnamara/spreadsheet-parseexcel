#!/usr/bin/perl
use strict;
use warnings;

use Test::More;
my $tests;
plan tests => $tests;
use Data::Dumper;

use_ok('Spreadsheet::ParseExcel');
BEGIN { $tests += 1; }


# these tests were created based on the values received using 0.27
# to create regressions tests

{
    # historically Parse returned and unblessed reference on missing file 
    my $excel = Spreadsheet::ParseExcel::Workbook->Parse('no_such_file.xls');
    is(ref($excel), 'HASH', 'failed Parse method returns HASH ref');
    isa_ok($excel->{_Excel}, 'Spreadsheet::ParseExcel', 
            'failed Parse method creates _Excel object');
    BEGIN { $tests += 2; }
}
{
    # historically Parse returned and unblessed reference on failure (e.g.
    # input file is not an Excel file)
    my $excel = Spreadsheet::ParseExcel::Workbook->Parse($0);
    is(ref($excel), 'HASH', 'failed Parse method returns HASH ref');
    isa_ok($excel->{_Excel}, 'Spreadsheet::ParseExcel', 
            'failed Parse method creates _Excel object');
    BEGIN { $tests += 2; }
}

my $workbook_1;
{
    my $workbook = Spreadsheet::ParseExcel::Workbook->Parse('sample/Excel/Test95.xls');
    $workbook_1 = $workbook;
    use Data::Dumper;
    #diag Dumper $excel;
    #_save_file('dump.txt', Dumper $excel);
    is(ref($workbook), 'Spreadsheet::ParseExcel::Workbook',
            'Spreadsheet::ParseExcel::Workbook created');
    my $excel = $workbook->{_Excel};
    isa_ok($excel, 'Spreadsheet::ParseExcel', 
            'Parse method creates _Excel object');
    is(ref($excel->{FuncTbl}), 'HASH');
    is(ref($excel->{GetContent}), 'CODE');


    # meta data
    is($workbook->{_CurSheet_}, 1, 'current sheet is 1');
    is($workbook->{_CurSheet}, 1);
    is($workbook->{Flg1904}, 0);
    isa_ok($workbook->{FmtClass}, 'Spreadsheet::ParseExcel::FmtDefault');
    
    # TODO more tests in Format
    is(ref($workbook->{Format}), 'ARRAY');
    my $formats = $workbook->{Format};
    is(scalar(@$formats), 22);
    # all but 2 are 'Spreadsheet::ParseExcel::Format' objects
    is(ref($workbook->{FormatStr}), 'HASH');
    is($workbook->{SheetCount}, 2);
    my $fonts = $workbook->{Font};
    is(ref($fonts), 'ARRAY');
    is(scalar(@$fonts), 6); 

    is($workbook->{Version}, 1280);
    is($workbook->{BIFFVersion}, 8);
    is($workbook->{File}, 'sample/Excel/Test95.xls');
    is($workbook->{Author}, 'kawait');



    my @sheets = @{$workbook->{Worksheet}};
    is (@sheets, 2, "two sheets");
    is($sheets[0]->{Name}, 'Sheet1-ASC');     # Open Office shows: 'Sheet1_ASC'
    is($sheets[1]->{Name}, 'Sheet1-ASC (2)'); # OO shows 'Sheet1_ASC_2_' 

    is($sheets[0]->{MinRow}, 0);
    is($sheets[0]->{MaxRow}, 7);
    #diag Dumper $sheets[0]->{Cells};
    #qw(ASC Date INTEGER Float Double Formula)
    is($sheets[0]->{Cells}[0][0]->{Val}, 'ASC');
    #diag Dumper $sheets[0]->{Cells}[0][0];


    is($sheets[1]->{MinRow}, 0);
    is($sheets[1]->{MaxRow}, 5);


    BEGIN { $tests += 26; }
}

eval "require IO::Scalar";
if ($@) {
    ok (1, "Skipped - no IO::Scalar") for 1..6;
    }
else {
{
    open my $fh, '<','t/excel_files/Test95.xls';
    my $workbook = Spreadsheet::ParseExcel::Workbook->Parse($fh);
    isnt($workbook, $workbook_1);
    delete $workbook_1->{File};  # when give a filehandlres this field is not set
    is_deeply($workbook, $workbook_1);
    BEGIN { $tests += 2; }
}

# pass a reference to a scalar containing the file content
{
    my $data;
    if (open my $fh, '<','t/excel_files/Test95.xls') {
        binmode($fh);
        local $/ = undef;
        $data = <$fh>;
    }
    my $workbook = Spreadsheet::ParseExcel::Workbook->Parse(\$data);
    isnt($workbook, $workbook_1);
    is_deeply($workbook, $workbook_1);
    BEGIN { $tests += 2; }
}
{
    open my $fh, '<','t/excel_files/Test95.xls';
    binmode($fh);
    my @data = <$fh>;
    my $workbook = Spreadsheet::ParseExcel::Workbook->Parse(\@data);
    isnt($workbook, $workbook_1);
    $workbook_1->{File} = undef;
    is_deeply($workbook, $workbook_1);
    BEGIN { $tests += 2; }
}
}

eval "require IO::Wrap";
if ($@) {
    ok (1, "Skipped - no IO::Wrap") for 1..4;
    }
else {
{
    open my $fh, '<','t/excel_files/Test95.xls';
    my $workbook = Spreadsheet::ParseExcel::Workbook->Parse($fh);
    isnt($workbook, $workbook_1);
    delete $workbook_1->{File};  # when give a filehandlres this field is not set
    is_deeply($workbook, $workbook_1);
    BEGIN { $tests += 2; }
}

# pass an IO::Wrap object
{
    my $data;
    my $fh;
    if (open my $real_fh, '<','t/excel_files/Test95.xls') {
        binmode($real_fh);
        $fh = IO::Wrap::wraphandle($real_fh);
    }
    my $workbook = Spreadsheet::ParseExcel::Workbook->Parse($fh);
    isnt($workbook, $workbook_1);
    is_deeply($workbook, $workbook_1);
    BEGIN { $tests += 2; }
}
}


sub _save_file {
    my ($file, $data) = @_;
    if (open my $fh, '>', $file) {
        print {$fh} $data;
    }
}
