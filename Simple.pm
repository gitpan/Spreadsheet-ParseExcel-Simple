package Spreadsheet::ParseExcel::Simple;

use strict;
use Spreadsheet::ParseExcel;
use vars qw/$VERSION/;
$VERSION = '0.01';

=head1 NAME

Spreadsheet::ParseExcel::Simple - A simple interface to Excel data

=head1 SYNOPSIS

  my $xls = Spreadsheet::ParseExcel::Simple->read('spreadsheet.xls');
  foreach my $sheet ($xls->sheets) {
     while ($sheet->has_data) {  
         my @data = $sheet->next_row;
     }
  }

=head1 DESCRIPTION

This provides an abstraction to the Spreadsheet::ParseExcel module for
simple reading of values.

You simply loop over the sheets, and fetch rows to arrays.

For anything more complex, you probably want to use
Spreadsheet::ParseExcel directly.

=head1 METHODS

=head2 read

  my $xls = Spreadsheet::ParseExcel::Simple->read('spreadsheet.xls');

This opens the spreadsheet specified for you. 

sub book { shift->{book} }

=head2 sheets

  @sheets = $xls->sheets;

Each spreadsheet can contain one or more worksheets. This fetches them
all back. You can then iterate over them, or jump straight to the one
you wish to play with.

=head2 has_data

  if ($sheet->has_data) { ... }

This lets us know if there are more rows in this sheet that we haven't
read yet. This allows us to differentiate between an empty row, and 
the end of the sheet.

=head2 next_row

  my @data = $sheet->next_row;

Fetch the next row of data back.

=head1 AUTHOR

Tony Bowden, E<lt>kasei@tmtm.comE<gt>.

=head1 SEE ALSO

L<Spreadsheet::ParseExcel>. 

=head1 COPYRIGHT

Copyright (C) 2001 Tony Bowden. All rights reserved.

This module is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut

sub read {
  my $class = shift;
  bless {
    book => Spreadsheet::ParseExcel->new->Parse(shift),
  }, $class;
}

sub sheets {
  my $self = shift;
  return map Spreadsheet::ParseExcel::Simple::_Sheet->new($_), 
                @{$self->{book}->{Worksheet}};
}

package Spreadsheet::ParseExcel::Simple::_Sheet;

sub new {
  my $class = shift;
  my $sheet = shift;
  bless {
    sheet => $sheet,
    row   => $sheet->{MinRow},
  }, $class;
}

sub has_data { $_[0]->{row} <= $_[0]->{sheet}->{MaxRow} } 

sub next_row {
  map { $_ ? $_->Value : "" } @{$_[0]->{sheet}->{Cells}[$_[0]->{row}++]};
}

1;

