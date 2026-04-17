#!/usr/bin/perl -w
use warnings;
use strict;

use Spreadsheet::Nifty::XLSB::RecordReader;

package Spreadsheet::Nifty::XLSB::BinaryIndex;

# The binary index contains groups of up to 32 rows that we're calling blocks.
# For each row that contains data, the row is divided into 16 groups of 1024
#  cells that we're calling spans.

use constant
{
  BI_MAX_ROWS_PER_BLOCK => 32,

  BI_CELLS_PER_SPAN     => 1024,
  BI_SPANS_PER_ROW      => 16,
};

# === Class methods ===

sub new()
{
  my $class = shift();

  my $self = {};
  $self->{debug}  = 0;
  $self->{blocks} = [];

  bless($self, $class);

  return $self;
}

# === Instance methods ===

# Reads the entirety of the index.
sub read($)
{
  my $self = shift();
  my ($member) = @_;

  my $block;

  my $recs = Spreadsheet::Nifty::XLSB::RecordReader->new($member);
  while (my $rec = $recs->read())
  {
    #printf("REC type %d size %d name %s\n", $rec->{type}, $rec->{size}, $rec->{name} // '?');

    if ($rec->{name} eq 'BrtIndexBlock')
    {
      defined($block) && die('Saw two BrtIndexBlock records in a row');

      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $brtIndexBlock = $decoder->decodeHash(['minRow:u32', 'nextRow:u32']);
      $brtIndexBlock->{count} = $brtIndexBlock->{nextRow} - $brtIndexBlock->{minRow};
      ($brtIndexBlock->{count} > BI_MAX_ROWS_PER_BLOCK) && die('BrtIndexBlock has row count above maximum');
      #print main::Dumper($brtIndexBlock);

      $block = $brtIndexBlock;
    }
    elsif ($rec->{name} eq 'BrtIndexRowBlock')
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $brtIndexRowBlock = $decoder->decodeHash(['rowMask:u32', 'offset:u64']);

      # Read column span masks for each row
      $brtIndexRowBlock->{spanMasks} = [];
      for (my $r = 0; $r < $block->{count}; $r++)
      {
        if ($brtIndexRowBlock->{rowMask} & (0x01 << $r))
        {
          my $colMask = $decoder->decodeField('u16');
          $brtIndexRowBlock->{spanMasks}->[$r] = $colMask;
        }
      }

      # Read indices for the start of each column span
      my $indices = [];
      my $maxIndex = undef;
      for (my $r = 0; $r < $block->{count}; $r++)
      {
        if ($brtIndexRowBlock->{rowMask} & (0x01 << $r))
        {
          my $spanMask = $brtIndexRowBlock->{spanMasks}->[$r];
          (!defined($spanMask)) && next;  # Don't bother if row has no spans

          for (my $s = 0; $s < BI_SPANS_PER_ROW; $s++)
          {
            if ($spanMask & (0x01 << $s))
            {
              $indices->[$r]->[$s] = $decoder->decodeField('u32');
            }
          }

          $maxIndex = $r;
        }
      }

      $block->{offset}  = $brtIndexRowBlock->{offset};
      $block->{indices} = $indices;
      $block->{maxIndex} = $maxIndex;
      push(@{$self->{blocks}}, $block);
      $block = undef;
    }
    elsif ($rec->{name} eq 'BrtIndexPartEnd')
    {
      last;
    }
  }

  #print main::Dumper($self->{blocks});

  return !!1;
}

# Get the offset of the last known cell span.
# This can be useful for reading sheet data that follows the cells.
sub getFinalOffset()
{
  my $self = shift();

  my $blockOffset = $self->{blocks}->[-1]->{offset};
  my $spanOffset = $self->{blocks}->[-1]->{indices}->[-1]->[-1];
  return $blockOffset + $spanOffset;
}

sub getRowCount()
{
  my $self = shift();

  (scalar(@{$self->{blocks}}) == 0) && return 0;  # No blocks means no rows

  # NOTE: The nextRow field is actually the lowest possible starting row index
  #  of the following block, so it could be beyond the end of the sheet.
  #return $self->{blocks}->[-1]->{nextRow};

  my $finalBlock = $self->{blocks}->[-1];
  (!defined($finalBlock->{maxIndex})) && die("Block contains no populated rows?");
  my $rowIndex = $finalBlock->{minRow} + $finalBlock->{maxIndex};

  return $rowIndex + 1;
}

# Returns the index block for a given row index.
# If no such block exists in the index, returns undef.
sub getBlockByRowIndex($)
{
  my $self = shift();
  my ($targetRow) = @_;

  my $blockCount = scalar(@{$self->{blocks}});
  #printf("getBlockByRowIndex(): blockCount %d  targetRow %d\n", $blockCount, $targetRow);

  # Don't bother with search if row index is out of range
  (($targetRow < 0) || ($targetRow > $self->{blocks}->[-1]->{nextRow})) && return undef;

  # Binary search to find the block
  my $low = 0;
  my $high = $blockCount;
  do
  {
    my $i = ($low + $high) >> 1;

    my $b = $self->{blocks}->[$i];
    #printf("  low %d  high %d  i %d  minRow %d\n", $low, $high, $i, $b->{minRow});

    if ($b->{minRow} <= $targetRow)
    {
      if ($b->{nextRow} > $targetRow)
      {
        #printf("    Found\n");
        return $b;
      }
      else
      {
        $low = $i + 1;
      }
    }
    else
    {
      $high = $i - 1;
    }
  } while ($low <= $high);

  return undef;  # Not found
}

1;
