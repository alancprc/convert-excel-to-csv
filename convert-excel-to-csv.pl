#!/usr/bin/env perl

use 5.010.001;
use strict;
use warnings;
use Function::Parameters qw(fun);
use Types::Standard qw(Str Int ArrayRef RegexpRef Object);
use Spreadsheet::ParseXLSX;
use Text::CSV_XS qw(csv);

=head1 description
 convert each sheet of a xlsx file to a csv file.
=cut

=head1 usage
 ./convert-excel-to-csv.pl sample.xlsx
=cut

sub main
{
    for my $excel (@ARGV) {
        ConvertExcelToCsvs($excel);
    }
}

fun ConvertExcelToCsvs (Str $excel)
{
    my $workbook = Spreadsheet::ParseXLSX->new->parse($excel);
    for my $worksheet ( $workbook->worksheets() ) {
        ConvertWorkSheetToCsv($worksheet);
    }
}

fun ConvertWorkSheetToCsv (Object $worksheet)
{
    my $sheetData;
    my ( $row_start, $col_start ) = ( 0, 0 );
    my $row_max = $worksheet->row_range();
    my $col_max = $worksheet->col_range();

    for my $row ( $row_start .. $row_max ) {
        for my $col ( $col_start .. $col_max ) {
            my $cell = $worksheet->get_cell( $row, $col );
            $sheetData->[$row][$col] = $cell ? $cell->unformatted() : undef;
        }
    }
    csv( in => $sheetData, out => $worksheet->get_name() . ".csv" );
}

&main();
