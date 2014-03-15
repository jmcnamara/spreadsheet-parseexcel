#!perl -w

use strict;
use Test::More;

eval "use Spreadsheet::WriteExcel; 1;"
	or plan( skip_all => "Spreadsheet::ParseExcel::SaveParser requires Spreadsheet::WriteExcel" );

use utf8;
use Encode qw(encode);

plan(tests => 7);

use_ok ('Spreadsheet::ParseExcel::SaveParser');

use_ok ('Spreadsheet::WriteExcel');

my $xl_base_name = 't/excel_files/46_save_parser.xls';

my $xl_base = Spreadsheet::WriteExcel->new($xl_base_name);

# testing merged cells

# first, we need to generate excel file with merged cells
my $worksheet = $xl_base->add_worksheet();

my $format = $xl_base->add_format(
	border  => 6,
	valign  => 'vcenter',
	align   => 'center',
);

$worksheet->merge_range('A1:B2', 'V & H', $format);

$worksheet->merge_range('E5:H8', 'V & H', $format);

$worksheet->fit_to_pages(1, 1);

$xl_base->close;

# parse excel and write modified file
my $xl_parser = Spreadsheet::ParseExcel::SaveParser->new;
my $template = $xl_parser->Parse($xl_base_name);

# test writing data to merged cell
$template->worksheet (0)->AddCell (4, 4, 'V & H mod');

my $workbook;
{
    local $^W = 0;
 
    $workbook = $template->SaveAs ($xl_base_name . '.mod.xls');
}

$workbook->close;

# parse modified file and check for merged cell

my $template_mod = $xl_parser->Parse ($xl_base_name . '.mod.xls');

my $worksheet_mod = $template_mod->worksheet (0);

my $merged_areas = $worksheet_mod->get_merged_areas;

ok scalar @$merged_areas == 2, 'merged areas count';

my @fit = $worksheet_mod->get_fit_to_pages;

is_deeply (\@fit, [1, 1], 'fix for fit to pages');

#use Data::Dumper;
#warn Dumper $merged_areas;

# RowHeight

is_deeply $merged_areas->[0], [0, 0, 1, 1];

is_deeply $merged_areas->[1], [4, 4, 7, 7], 'overwritten merged cell position';

ok $worksheet_mod->Cell (4, 4)->value eq 'V & H mod', 'overwritten merged cell value';

unlink $xl_base_name;

unlink $xl_base_name . '.mod.xls';

1;
