#!/usr/bin/perl -w

use strict;
use Spreadsheet::ParseExcel::Simple;
use Test::More tests => 11;

eval {
  require File::Temp;
  require Spreadsheet::WriteExcel::Simple;   
};

if ($@) {
  ok(1, "Need File::Temp and Spreadsheet::WriteExcel::Simple");
  ok(1, " for sensible testing.");
  ok(1, "  - skipping tests") for (1 .. 6);
} else {
  File::Temp->import(qw/tempfile tempdir/);  
  my $dir1 = tempdir(CLEANUP => 1);
  my ($fh1, $name1) = tempfile(DIR => $dir1); 
    
  my @row1 = qw/foo bar baz/;
  my @row2 = qw/1 fred 2001-01-01/;
  my @row3 = ();
  my @row4 = (2, undef, "2001-03-01");

  # Write our our test file.
  my $ss = Spreadsheet::WriteExcel::Simple->new;
     $ss->write_bold_row(\@row1);
     $ss->write_row(\@row2);
     $ss->write_row(\@row3);
     $ss->write_row(\@row4);
  print $fh1 $ss->data;
  close $fh1;

  # Now read it back in
  my $xls = Spreadsheet::ParseExcel::Simple->read($name1);
  my @sheets = $xls->sheets;
  is scalar @sheets, 1, "We have one sheet";
  my $sheet = $sheets[0];

  ok $sheet->has_data, "We have data to read";
  my @fetch1 = $sheet->next_row;
  ok eq_array(\@fetch1, \@row1), "Header OK";

  ok $sheet->has_data, "We still have data to read";
  my @fetch2 = $sheet->next_row;
  ok eq_array(\@fetch2, \@row2), "Row 2";

  ok $sheet->has_data, "We still have data to read";
  my @fetch3 = $sheet->next_row;
  ok eq_array(\@fetch3, \@row3), "Row 3 (blank)";

  ok $sheet->has_data, "We still have data to read";
  my @fetch4 = $sheet->next_row;
  ok eq_array(\@fetch4, \@row4), "Row 4";

  ok !$sheet->has_data, "No more data to read";
  ok !$sheet->next_row, "So, can't read any";
}
