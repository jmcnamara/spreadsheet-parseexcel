#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::ParseExcel.
#
# Test for get_active_sheet
#

use strict;

use Test::More tests => 45;

use Spreadsheet::ParseExcel;

##############################################################################
#
# Tests.
#

my $parser = Spreadsheet::ParseExcel->new;

# Workbook saved with sheet2 (index 1) open

my $book = $parser->parse( "t/excel_files/TestActiveSheet.xls" );
my $active = $book->get_active_sheet;
is($active, 1, 'Correct sheet');

my $ws = $book->worksheet('Sheet2');

my %expect = (
              A3 => { },
              # Each entry adds 6 tests
              A6 => { desc => q(http://www.example.com),
                      link => q(http://www.example.com/),
                    },
              B6 => { desc => q(http://www.example.com#foo),
                      link => q(http://www.example.com/#foo),
                    },
              C6 => { desc => q(file:///c:\\nodir\\nofile.txt),
                      link => q(file:///c:\\nodir\\nofile.txt),
                    },
              D6 => { desc => q(\\\\server\\quirks\\sometest.bat),
                      link => q(file:///\\\\server\\quirks\\sometest.bat),
                    },
              E6 => { desc => q(TestActiveSheet.xls),
                      rel  => 1,
                      link => q(TestActiveSheet.xls),
                    },
              F6 => { desc => q(Sheet2!A7),
                      link => q(#Sheet2%21A7),
                    },
              A7 => { desc => q(www.example.com),
                      link => q(http://www.example.com/),
                    },
              B7 => { desc => q(www.example.com#foo),
                      link => q(http://www.example.com/#foo),
                    },
              C7 => { desc => q(c:\\nodir\\nofile.txt),
                      link => q(file:///c:\\nodir\\nofile.txt),
                    },
              D7 => { desc => q(SMB Link Sometest.bat),
                      link => q(file:///\\\\server\\quirks\\sometest.bat),
                    },
              E7 => { desc => q(Rel: TestActiveSheet.xls),
                      rel  => 1,
                      link => q(TestActiveSheet.xls),
                    },
              F7 => { desc => q(mailto:fred@example.net),
                      link => q(mailto:fred@example.net),
                    },
              A9 => { desc => q(file:///..\\..\\zipple.dat),
                      link => q(../../zipple.dat),
                    },
              A10 => { desc => q(ftp://user:pass@example.net/pub/manuals/Excel.doc),
                       link => q(ftp://user:pass@example.net/pub/manuals/Excel.doc),
                    },
              );

foreach my $t (sort keys %expect) {
    my $link = $expect{$t}{link};
    my $desc = $expect{$t}{desc};
    my $rel = $expect{$t}{rel};
    if( $rel ) {
        $link = "file:///t/excel_files/$link";
    }

    $t =~ m/^(.)(\d+)$/ or die;
    my $cell = $ws->get_cell( $2-1, (ord($1)-ord('A')) );
    ok( defined $cell, "Cell $t defined" );
    next unless defined $cell;

    my $hl = $cell->get_hyperlink;
    if( !defined $link ) {
        is( $hl, undef, "Cell $t should have no link" );
        next;
    }
    ok( defined $hl->[0] && $hl->[0] eq $desc, "Cell $t description match" );
    ok( defined $hl->[1] && $hl->[1] eq $link, "Cell $t link match" );
}



__END__
