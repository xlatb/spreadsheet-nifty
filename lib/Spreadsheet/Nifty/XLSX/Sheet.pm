#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSX::Sheet;

use Spreadsheet::Nifty::XLSX::Cell;

# === Class methods ===

sub new()
{
  my $class = shift();
  my ($reader) = @_;
  
  my $self = {};
  $self->{reader}   = $reader;
  $self->{rowIndex} = 0;
  
  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub seekRow($)
{
  my $self = shift();
  my ($rowIndex) = @_;

  $self->{rowIndex} = $rowIndex;
  return;
}

sub tellRow()
{
  my $self = shift();

  return $self->{rowIndex};
}

sub getName()
{
  my $self = shift();

  return $self->{reader}->{workbook}->getSheetNames()->[$self->{reader}->{sheetIndex}];
}

# TODO: Dimensions could be missing and we'd need to scan
sub getRowDimensions()
{
  my $self = shift();

  (!defined($self->{reader}->{header}->{dimensions})) && return undef;  # Unknown dimensions

  return [$self->{reader}->{header}->{dimensions}->{minRow}, $self->{reader}->{header}->{dimensions}->{maxRow}];
}

# TODO: Dimensions could be missing and we'd need to scan
sub getColDimensions()
{
  my $self = shift();

  (!defined($self->{reader}->{header}->{dimensions})) && return undef;  # Unknown dimensions

  return [$self->{reader}->{header}->{dimensions}->{minCol}, $self->{reader}->{header}->{dimensions}->{maxCol}];
}

sub buildCell($)
{
  my $self = shift();
  my ($data) = @_;

  (!defined($data)) && return undef;

  my $type = $data->{type};
  my $value = $data->{value};

  if ($type eq 'n')
  {
    if (!defined($value))
    {
      return Spreadsheet::Nifty::XLSX::Cell->new(Spreadsheet::Nifty::TYPE_NULL);
    }

    return Spreadsheet::Nifty::XLSX::Cell->new(Spreadsheet::Nifty::TYPE_NUM, $value);
  }
  elsif ($type eq 'inlineStr')  # Inline string
  {
    return Spreadsheet::Nifty::XLSX::Cell->new(Spreadsheet::Nifty::TYPE_STR, $value);
  }
  elsif ($type eq 's')  # Shared string
  {
    return Spreadsheet::Nifty::XLSX::Cell->new(Spreadsheet::Nifty::TYPE_STR, $self->{reader}->{workbook}->getSharedString($data->{stringIndex}));
  }
  elsif ($type eq 'str')  # Formula whose returned value is a string
  {
    return Spreadsheet::Nifty::XLSX::Cell->new(Spreadsheet::Nifty::TYPE_STR, $value);
  }
  elsif ($type eq 'b')
  {
    return Spreadsheet::Nifty::XLSX::Cell->new(Spreadsheet::Nifty::TYPE_BOOL, $value);
  }
  elsif ($type eq 'e')
  {
    return Spreadsheet::Nifty::XLSX::Cell->new(Spreadsheet::Nifty::TYPE_ERR, Spreadsheet::Nifty->errorNumber($value));
  }

  die("Unhandled cell type '$type'");
}

# Reads cells for the current row.
sub readRow()
{
  my $self = shift();

  my $row = $self->{reader}->findRow($self->{rowIndex});
  if (!defined($row))
  {
    # No such row in the file, so return a blank row
    $self->{rowIndex}++;
    return [];
  }
  elsif ($row == 0)
  {
    # Past end of file
    return undef;
  }

  $self->{rowIndex}++;
  my $cells = $self->{reader}->readRow();
  #print main::Dumper($cells);
  return [ map( { $self->buildCell($_) } @{$cells} ) ];
}

1;
