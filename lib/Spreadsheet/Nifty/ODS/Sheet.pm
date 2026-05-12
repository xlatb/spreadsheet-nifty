#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::ODS::Sheet;
use Spreadsheet::Nifty::ODS::Cell;
use Spreadsheet::Nifty::Utils;

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
    $cell = Spreadsheet::Nifty::ODS::Cell->new(Spreadsheet::Nifty::TYPE_NULL, undef, $self, $cellDef);
  }
  elsif (($valueType eq 'float') || ($valueType eq 'percentage'))
  {
    $cell = Spreadsheet::Nifty::ODS::Cell->new(Spreadsheet::Nifty::TYPE_NUM, $cellDef->{value}, $self, $cellDef);
  }
  elsif ($valueType eq 'string')
  {
    $cell = Spreadsheet::Nifty::ODS::Cell->new(Spreadsheet::Nifty::TYPE_STR, $cellDef->{value}, $self, $cellDef);
  }
  elsif ($cellDef->{valueType} eq 'boolean')
  {
    $cell = Spreadsheet::Nifty::Cell->new(Spreadsheet::Nifty::TYPE_BOOL, Spreadsheet::Nifty::ODS::Decode->decodeBoolean($cellDef->{value}), $self, $cellDef);
  }
  elsif ($valueType eq 'date')
  {
    my $struct = Spreadsheet::Nifty::ODS::Decode->decodeDateString($cellDef->{value});
    my $value = Spreadsheet::Nifty::Utils->structToExcelTime($struct);
    $cell = Spreadsheet::Nifty::ODS::Cell->new(Spreadsheet::Nifty::TYPE_DATE, $value, $self, $cellDef);
  }
  else
  {
    use Data::Dumper;
    print Dumper($cellDef);
    die("Unimplemented valueType '$cellDef->{valueType}'");
    ...;
  }

  return $cell;
}

# Reads cells for the current row.
sub readRow()
{
  my $self = shift();

  my $defs = $self->{sheetReader}->readRowAsDefs();
  (!defined($defs)) && return undef;

  my $cells = [];
  my $x = 0;
  for my $def (@{$defs->{cellDefs}})
  {
    # Handle empty cells
    if ($def->{empty})
    {
      $x += $def->{count};
      next;
    }

    my $cell = $self->buildCell($def);

    # Apply default style if the cell def did not have one
    if (!defined($def->{style}))
    {
      $cell->{p}->{style} = $defs->{rowDef}->{cellStyle} // $self->{sheetReader}->{columnDefs}->{defs}->[$self->{sheetReader}->{columnDefs}->{byIndex}->[$x]]->{cellStyle};
    }

    $cells->[$x++] = $cell;

    my $count = $def->{count};
    if ($count > 1)
    {
      for (my $i = 2; $i <= $count; $i++)
      {
        my $dup = $cell->dup();

        # Apply default style if the cell def did not have one
        if (!defined($def->{style}))
        {
          $dup->{p}->{style} = $defs->{rowDef}->{cellStyle} // $self->{sheetReader}->{columnDefs}->{defs}->[$self->{sheetReader}->{columnDefs}->{byIndex}->[$x]]->{cellStyle};
        }

        $cells->[$x++] = $dup;
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

  return $self->{sheetReader}->seekRow($rowIndex);
}

sub tellRow()
{
  my $self = shift();

  return $self->{sheetReader}->tellRow();
}

1;
