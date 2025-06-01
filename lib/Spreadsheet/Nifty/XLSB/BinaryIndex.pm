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
        }
      }

      $block->{offset}  = $brtIndexRowBlock->{offset};
      $block->{indices} = $indices;
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

1;
