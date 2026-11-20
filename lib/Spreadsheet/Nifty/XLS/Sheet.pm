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
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_NULL, undef, $self->{reader}, $cell);
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::LABEL)
    {
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_STR, $cell->{str}, $self->{reader}, $cell);
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::LABEL_SST)
    {
      my $str = $self->{workbook}->{workbook}->{strings}->[$cell->{si}]->{str};
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_STR, $str, $self->{reader}, $cell);
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::RSTRING)
    {
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_STR, $cell->{str}, $self->{reader}, $cell);
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::RK)
    {
      my $num = Spreadsheet::Nifty::Utils->translateRk($cell->{rk});
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_NUM, $num, $self->{reader}, $cell);
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::NUMBER)
    {
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_NUM, $cell->{num}, $self->{reader}, $cell);
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::FORMULA)
    {
      my $data = $self->decodeFormulaResult($cell);
      $data->{f} = $cell->{formula};
      $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new($data->{t}, $data->{v}, $self->{reader}, $cell);
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::BOOL_ERR)
    {
      if ($cell->{datatype} == 0)
      {
        $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_BOOL, $cell->{value}, $self->{reader}, $cell);
      }
      elsif ($cell->{datatype} == 1)
      {
        $row->[$cell->{col}] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_ERR, $cell->{value}, $self->{reader}, $cell);
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
        $row->[$cell->{minCol} + $i] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_NUM, $v, $self->{reader}, $cell);
      }
    }
    elsif ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::MUL_BLANK)
    {
      for (my $i = 0; $i < scalar(@{$cell->{xfs}}); $i++)
      {
        $row->[$cell->{minCol} + $i] = Spreadsheet::Nifty::XLS::Cell->new(Spreadsheet::Nifty::TYPE_NULL, undef, $self->{reader}, $cell);
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

# NOTE: If minRow == maxRow, the autofilter area's final row is not explicitly
#  defined. The area will be deemed to stop at the next empty row.
sub getAutoFilter()
{
  my $self = shift();

  my $label = $self->{workbook}->getLabelByNameAndWorksheet('_FilterDatabase', $self->{sheetIndex});
  (!defined($label)) && return undef;

  my $decoder = $self->{reader}->getFormulaDecoder();
  my $formula = $decoder->decode($label->{formula});
  #print main::Dumper('UNPARSE', $formula->unparse({sheetReader => $self->{reader}}));
  ($formula->getTokenCount() != 1) && die("getAutoFilter(): Expected _FilterDatabase formula to have a single token");

  # The formula should have decoded to a single tArea3d token.
  my $token = $formula->getToken(0);
  ($token->{type} != 0x3B) && die("getAutoFilter(): Unexpected token type");

  # We could follow the token's external reference using something like:
  #  $self->{reader}->getExternalSheetReference($token->{externIndex})
  # But, this external link should always be pointing to a 'workbook-self'
  #  reference with minSheet and maxSheet set to this sheet.

  my $autoFilter = {range => {}};
  $autoFilter->{range}->{minRow} = $token->{ref}->{minRow};
  $autoFilter->{range}->{maxRow} = $token->{ref}->{maxRow};
  $autoFilter->{range}->{minCol} = $token->{ref}->{minCol};
  $autoFilter->{range}->{maxCol} = $token->{ref}->{maxCol};

  return $autoFilter;
}

1;
