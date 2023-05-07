#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLS::Sheet;

use Spreadsheet::Nifty::XLS::Cell;

# === Class methods ===

sub new()
{
  my $class = shift();
  my ($workbook, $sheetIndex, $reader) = @_;

  my $self = {};
  $self->{workbook}   = $workbook;
  $self->{sheetIndex} = $sheetIndex;
  $self->{reader}     = $reader;
  $self->{rowIndex}   = 0;

  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub decodeFormulaResult($)
{
  my $self = shift();
  my ($formula) = @_;

  # As a space-saving trick, if the last two bytes are not 0xFFFF, the entire
  #  8 bytes are an IEEE double. This works because all ones in those bit
  #  positions in a little-endian double would signify infinity or nan, which
  #  are not allowed values.
  if (unpack('v', substr($formula->{result}, 6, 2)) != 0xFFFF)
  {
    return {t => Spreadsheet::Nifty::TYPE_NUM, v => unpack('d<', $formula->{result})};
  }

  my $type = ord(substr($formula->{result}, 0, 1));
  if ($type == 0)  # String
  {
    return {t => Spreadsheet::Nifty::TYPE_STR, v => $formula->{string}};
  }
  elsif ($type == 1)  # Boolean
  {
    return {t => Spreadsheet::Nifty::TYPE_BOOL, v => ord(substr($formula->{result}, 2, 1))};
  }
  elsif ($type == 2)  # Error
  {
    return {t => Spreadsheet::Nifty::TYPE_ERR, v => ord(substr($formula->{result}, 2, 1))};
  }
  elsif ($type == 3)  # Blank string
  {
    return {t => Spreadsheet::Nifty::TYPE_STR, v => ''};
  }

  die("Unhandled formula result type $type");
}

# Reads cells for the current row.
sub readRow()
{
  my $self = shift();

  ($self->{rowIndex} >= $self->{reader}->{header}->{dimensions}->{maxRow}) && return undef;  # End of rows

  my $cells = $self->{reader}->readRowCells($self->{rowIndex});
  $self->{rowIndex}++;

  my $row = [];
  for my $cell (@{$cells})
  {
    if ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::BLANK)
    {
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_NULL);
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::LABEL)
    {
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_STR, $cell->{str});
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::LABEL_SST)
    {
      my $str = $self->{workbook}->{workbook}->{strings}->[$cell->{si}]->{str};
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_STR, $str);
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::RSTRING)
    {
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_STR, $cell->{str});
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::RK)
    {
      my $num = Spreadsheet::Nifty::Utils->translateRk($cell->{rk});
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_NUM, $num);
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::NUMBER)
    {
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_NUM, $cell->{num});
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::FORMULA)
    {
      my $data = $self->decodeFormulaResult($cell);
      $data->{f} = $cell->{formula};
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new($data->{t}, $data->{v});
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::BOOL_ERR)
    {
      if ($cell->{datatype} == 0)
      {
        $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_BOOL, $cell->{value});
      }
      elsif ($cell->{datatype} == 1)
      {
        $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_ERR, $cell->{value});
      }
      else
      {
        die("Unhandled BoolErr datatype");
      }
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::MUL_RK)
    {
      for (my $i = 0; $i < scalar(@{$cell->{recs}}); $i++)
      {
        my $v = Spreadsheet::Nifty::Utils->translateRk($cell->{recs}->[$i]->{rk});
        $row->[$cell->{minCol} + $i] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_NUM, $v);
      }
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::MUL_BLANK)
    {
      for (my $i = 0; $i < scalar(@{$cell->{xfs}}); $i++)
      {
        $row->[$cell->{minCol} + $i] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_NULL);
      }
    }
    else
    {
      die(sprintf("Unhandled cell type 0x%04X", $cell->{type}));
    }
  }
  
  return $row;
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

  return $self->{rowIndex};
}

sub getName()
{
  my $self = shift();

  return $self->{workbook}->getSheetNames()->[$self->{sheetIndex}];
}

sub getRowDimensions()
{
  my $self = shift();

  return [$self->{reader}->{header}->{dimensions}->{minRow}, $self->{reader}->{header}->{dimensions}->{maxRow} - 1];
}

sub getColDimensions()
{
  my $self = shift();

  return [$self->{reader}->{header}->{dimensions}->{minCol}, $self->{reader}->{header}->{dimensions}->{maxCol} - 1];
}

1;
