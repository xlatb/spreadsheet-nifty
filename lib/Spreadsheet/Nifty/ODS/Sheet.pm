#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::ODS::Sheet;
use Spreadsheet::Nifty::ODS::Cell;

# === Class methods ===

sub new()
{
  my $class = shift();
  my ($fileReader, $sheetIndex) = @_;
  
  my $self = {};
  $self->{fileReader}  = $fileReader;
  $self->{sheetIndex}  = $sheetIndex;
  $self->{sheetReader} = Spreadsheet::Nifty::ODS::SheetReader->new($fileReader, $sheetIndex);
  
  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub open()
{
  my $self = shift();

  return $self->{sheetReader}->open();
}

sub getName()
{
  my $self = shift();

  $self->{fileReader}->getSheetNames()->[$self->{sheetReader}->{sheetIndex}];
}

sub getRowDimensions()
{
  my $self = shift();

  return [0, $self->{fileReader}->getSheetRowCount($self->{sheetIndex})];
}

sub getColDimensions()
{
  my $self = shift();

  return [0, $self->{fileReader}->getSheetColCount($self->{sheetIndex})];
}

# Given a cellDef, returns a Cell object.
# The cellDef's repeat count is unused.
sub buildCell($)
{
  my $self = shift();
  my ($cellDef) = @_;

  my $valueType = $cellDef->{valueType};

  my $cell;
  if ($valueType eq 'void')
  {
    $cell = Spreadsheet::Nifty::ODS::Cell->new(Spreadsheet::Nifty::TYPE_NULL);
  }
  elsif (($valueType eq 'float') || ($valueType eq 'percentage'))
  {
    $cell = Spreadsheet::Nifty::ODS::Cell->new(Spreadsheet::Nifty::TYPE_NUM, $cellDef->{value});
  }
  elsif ($valueType eq 'string')
  {
    $cell = Spreadsheet::Nifty::ODS::Cell->new(Spreadsheet::Nifty::TYPE_STR, $cellDef->{value});
  }
#  elsif ($cellDef->{valueType} eq 'boolean')
#  {
#    $cell = Spreadsheet::Nifty::Cell->new(Spreadsheet::Nifty::TYPE_BOOL, $cellDef->{value});
#  }
  else
  {
    print main::Dumper($cellDef);
    die("Unimplemented valueType '$cellDef->{valueType}'");
    ...;
  }

  return $cell;
}

# Reads cells for the current row.
sub readRow()
{
  my $self = shift();

  my $cellDefs = $self->{sheetReader}->readRow();
  (!defined($cellDefs)) && return undef;

  my $cells = [];
  my $x = 0;
  for my $def (@{$cellDefs})
  {
    # Handle empty cells
    if ($def->{empty})
    {
      $x += $def->{count};
      next;
    }

    my $cell = $self->buildCell($def);
    $cells->[$x++] = $cell;

    my $count = $def->{count};
    if ($count > 1)
    {
      $x += $count - 1;
      for (my $i = 0; $i <= $count; $i++)
      {
        push(@{$cells}, $cell->dup());
      }
    }
  }
  
  #print main::Dumper($cells);
  return $cells;
}

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

  return $self->{sheetReader}->tellRow();
}

1;
