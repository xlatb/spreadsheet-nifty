#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLS::SheetReader;

use Data::Dumper;

# === Class methods ===

sub new()
{
  my $class = shift();
  my ($workbook, $index, $offset, $biff) = @_;

  my $self = {};
  $self->{workbook} = $workbook;
  $self->{index}    = $index;
  $self->{biff}     = $biff;
  $self->{header}   = undef;
  $self->{offsets}  = {bof => $offset, rows => [], firstCells => []};
  $self->{lengths}  = {rows => []};
  $self->{cache}    = {rowBlocks => []};
  $self->{meta}     = {};
  $self->{debug}    = $workbook->{debug};

  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub open()
{
  my $self = shift();

  my $success = $self->readHeader();
  (!$success) && return 0;

  $self->checkIndex();

  return $success;
}

sub getBiffVersion()
{
  my $self = shift();

  # NOTE: Allegedly some versions of Excel write the wrong BIFF version to
  #  sheet streams. So we'll trust the workbook BIFF version over the sheet
  #  version.
  return $self->{workbook}->{workbook}->{bof}->{version};
}

# Checks the file's index against the row blocks.
sub checkIndex()
{
  my $self = shift();

  (!defined($self->{header}->{index})) && return;  # No index to check

  # Each row block covers 32 rows, and we should start counting the rows
  #  beginning at the minRow field of the Index record. Unforunately, some
  #  files start counting at row zero regardless of the minRow field. Here
  #  we detect a sheet with these misaligned row blocks.

  # If the index starts with row zero, there would be no discrepancy.
  ($self->{header}->{index}->{minRow} == 0) && return;

  # Test row blocks. In the worst case we might need to scan every row block
  #  of the sheet, but this is highly unlikely.
  for (my $i = 0; $i < scalar(@{$self->{header}->{index}->{dbcellOffsets}}); $i++)
  {
    my $blockInfo = $self->readRowBlockInfo($i);
    (!defined($blockInfo)) && next;  # Couldn't read row block?
    (scalar(@{$blockInfo->{rowIndices}}) == 0) && next;  # No rows in block?

    my $firstRowIndex = $blockInfo->{rowIndices}->[0];
    my $rowIndexCount = scalar(@{$blockInfo->{rowIndices}});
    my $firstRowOffset = $firstRowIndex % Spreadsheet::Nifty::XLS::ROW_BLOCK_SIZE;
    my $lowestPossibleRowIndex = $firstRowIndex - (Spreadsheet::Nifty::XLS::ROW_BLOCK_SIZE - $rowIndexCount);  # Lowest possible row index if the row block was full (contained 32 rows with data)
    my $lowestPossibleRowOffset = $lowestPossibleRowIndex % Spreadsheet::Nifty::XLS::ROW_BLOCK_SIZE;
    ($self->{debug}) && printf("checkIndex(): block %d firstRowIndex %d firstRowOffset %d lowestPossibleRowIndex %d lowestPossibleRowOffset %d rowIndexCount %d\n", $i, $firstRowIndex, $firstRowOffset, $lowestPossibleRowIndex, $lowestPossibleRowOffset, $rowIndexCount);

    if ($firstRowOffset < $self->{header}->{index}->{minRow})
    {
      # Row blocks are probably counted from row zero
      ($self->{debug}) && printf("checkIndex(): enabling workaround for misaligned row blocks\n");
      $self->{meta}->{index}->{minRow} = 0;
      return;
    }
    elsif ($lowestPossibleRowOffset >= $self->{header}->{index}->{minRow})
    {
      ($self->{debug}) && printf("checkIndex(): Blocks seem to be aligned correctly\n");
      return;
    }
  }

  return;
}

# Reads a sheet's records up until the cell table (everything up until Dimensions record).
sub readHeader()
{
  my $self = shift();

  $self->{header} = {};
  $self->{header}->{colInfos} = [];

  $self->{biff}->seek($self->{offsets}->{bof});
  my $rec = $self->{biff}->readRecord();
  (!defined($rec)) && return 0;
  ($rec->{type} != Spreadsheet::Nifty::XLS::RecordTypes::BOF) && return 0;
  
  $self->{header}->{bof} = Spreadsheet::Nifty::XLS::Decode::decodeBOF($rec->{payload});

  my $biffVersion = $self->getBiffVersion();

  while (defined($rec = $self->{biff}->readRecord()))
  {
    if ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::EOF)
    {
      $self->{workbook}->recordSheetOffsets($self->{index}, {'eof' => $rec->{offset}, 'next' => $self->{biff}->tell()});
      last;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::INDEX)
    {
      $self->{header}->{index} = Spreadsheet::Nifty::XLS::Decode::decodeIndex($rec->{payload}, $biffVersion);
    }
    #elsif ($rec->{name} eq 'DefaultRowHeight')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::DEFAULT_ROW_HEIGHT)
    {
      $self->{header}->{defaultRowHeight} = Spreadsheet::Nifty::XLS::Decode::decodeDefaultRowHeight($rec->{payload});
    }
    #elsif ($rec->{name} eq 'WsBool')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::WS_BOOL)
    {
      $self->{header}->{wsBool} = Spreadsheet::Nifty::XLS::Decode::decodeWsBool($rec->{payload});
    }
    #elsif ($rec->{name} eq 'Password')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::PASSWORD)
    {
      $self->{header}->{protection}->{password} = unpack('v', $rec->{payload});
    }
    #elsif ($rec->{name} eq 'DefColWidth')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::DEF_COL_WIDTH)
    {
      $self->{header}->{defaultColWidth} = unpack('v', $rec->{payload});
    }
    #elsif ($rec->{name} eq 'ColInfo')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::COL_INFO)
    {
      push(@{$self->{header}->{colInfos}}, Spreadsheet::Nifty::XLS::Decode::decodeColInfo($rec->{payload}));
    }
    #elsif ($rec->{name} eq 'AutoFilterInfo')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::AUTO_FILTER_INFO)
    {
      $self->readAutoFilterSection($rec);
    }
    #elsif ($rec->{name} eq 'Dimensions')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::DIMENSIONS)
    {
      my $dimensions = Spreadsheet::Nifty::XLS::Decode::decodeDimensions($rec->{payload});

      # There are files that advertise more rows than possible
      if ($dimensions->{maxRow} > Spreadsheet::Nifty::XLS::MAX_ROW_COUNT)
      {
        $dimensions->{maxRow} = Spreadsheet::Nifty::XLS::MAX_ROW_COUNT;
      }

      $self->{header}->{dimensions} = $dimensions;
      $self->{offsets}->{body} = $self->{biff}->tell();
      last;
    }
  }

  return 1;
}

# Given an AutoFilterInfo record, reads the section.
sub readAutoFilterSection($)
{
  my $self = shift();
  my ($rec) = @_;

  ($rec->{name} eq 'AutoFilterInfo') || die("readAutoFilterSection(): Expected AutoFilterInfo record");

  my $count = unpack('v', $rec->{payload});

  $self->{header}->{autoFilter} = {};
  $self->{header}->{autoFilter}->{colCount} = $count;
  $self->{header}->{autoFilter}->{entries} = [];

  my $lastOffset = $self->{biff}->tell();

  # Following records are one per *filtered* column.
  while (defined($rec = $self->{biff}->readRecord()))
  {
    ($self->{debug}) && printf("AutoFilter rec: name '%s' length %d\n", $rec->{name}, $rec->{length});
    #if ($rec->{name} eq 'AutoFilter')
    if ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::AUTO_FILTER)
    {
      #printf("AutoFilter\n");
      push(@{$self->{header}->{autoFilter}->{entries}}, Spreadsheet::Nifty::XLS::Decode::decodeAutoFilter($rec->{payload}));
      $lastOffset = $self->{biff}->tell();
    }
    #elsif ($rec->{name} eq 'AutoFilter12')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::AUTO_FILTER_12)
    {
      # TODO: Decode?
      #push(@{$self->{header}->{autoFilter}->{entries}}, Spreadsheet::Nifty::XLS::Decode::decodeAutoFilter12($rec->{payload}));
      #$lastOffset = $self->{biff}->tell();
    }
    else
    {
      # Not an Autofilter entry, so back up
      $self->{biff}->seek($lastOffset);
      last;
    }
  }

  return 1;
}

# Given a row index, returns associated indices:
# {blockIndex, minRowIndex, maxRowIndex}
# Return undef if the row is outside of any block.
sub getRowBlockIndices($)
{
  my $self = shift();
  my ($rowIndex) = @_;

  my $index = $self->{header}->{index};
  (!defined($index)) && return undef;  # No Index record

  my $minRow = $self->{meta}->{index}->{minRow} // $index->{minRow};

  # Bounds check
  ($rowIndex < $minRow) && return undef;
  ($rowIndex >= $index->{maxRow}) && return undef;  # NOTE: maxRow is really the first unused row

  my $blockIndex = int(($rowIndex - $minRow) / Spreadsheet::Nifty::XLS::ROW_BLOCK_SIZE);
  my $minRowIndex = $minRow + ($blockIndex * Spreadsheet::Nifty::XLS::ROW_BLOCK_SIZE);
  my $maxRowIndex = $minRowIndex + Spreadsheet::Nifty::XLS::ROW_BLOCK_SIZE - 1;
  #printf("Expecting row block %d to have Row records with indices within %d..%d\n", $blockIndex, $minRowIndex, $maxRowIndex);

  return {blockIndex => $blockIndex, minRowIndex => $minRowIndex, maxRowIndex => $maxRowIndex};
}

# Given a row index, returns information about the row block:
# {dbcell, rowIndices, nextOffset}
sub readRowBlockInfo($)
{
  my $self = shift();
  my ($blockIndex) = @_;

  # If we have this block cached, we can return it directly
  if (defined($self->{cache}->{rowBlocks}->[$blockIndex]))
  {
    return $self->{cache}->{rowBlocks}->[$blockIndex];
  }

  my $index = $self->{header}->{index};
  (!defined($index)) && return undef;  # No Index record

  # Block bounds check
  ($blockIndex < 0) && return undef;
  ($blockIndex >= scalar(@{$index->{dbcellOffsets}})) && return undef;  # Past end of known row blocks

  # Try to read DBCell record for this block
  my $dbcellOffset = $index->{dbcellOffsets}->[$blockIndex];
  ($self->{debug}) && printf("readRowBlockInfo(): blockIndex: %d dbcellOffset: %d\n", $blockIndex, $dbcellOffset);
  $self->{biff}->seek($dbcellOffset);
  my $rec = $self->{biff}->readRecord();
  (!defined($rec)) && return undef;  # End of records
  (($rec->{name} // '') ne 'DBCell') && return undef;  # Wrong record type

  # Decode DBCell and adjust the firstRowOffset from relative to absolute
  my $dbcell = Spreadsheet::Nifty::XLS::Decode::decodeDBCell($rec->{payload});
  $dbcell->{firstRowOffset} = $rec->{offset} - $dbcell->{firstRowOffset};

  my $blockInfo = {dbcell => $dbcell, rowIndices => []};
  $self->{cache}->{rowBlocks}->[$blockIndex] = $blockInfo;
  ($self->{debug}) && printf("readRowBlockInfo(): firstRowOffset: %d\n", $dbcell->{firstRowOffset});

  # Read this block's Row records
  ($self->{debug}) && printf("readRowBlockInfo(): Reading Row records...\n");
#  my $seenRows = {};
  my $rowIndices = [];
  $self->{biff}->seek($blockInfo->{dbcell}->{firstRowOffset});
  while (defined(my $rec = $self->{biff}->readRecord()))
  {
    if (scalar(@{$rowIndices}) >= Spreadsheet::Nifty::XLS::ROW_BLOCK_SIZE)
    {
      # Reached maxiumum block size
      $blockInfo->{nextOffset} = $self->{biff}->tell();
      last;
    }

    if (($rec->{name} // '') ne 'Row')
    {
      # Out of Row records in this block
      $blockInfo->{nextOffset} = $rec->{offset};
      last;
    }

#    my $row = Spreadsheet::Nifty::XLS::Decode::decodeRow($rec->{payload});
#    ($self->{debug}) && print main::Dumper('ROW', $row);
#
#    if (!defined($self->{offsets}->{rows}->[$row->{row}]))
#    {
#      $self->{offsets}->{rows}->[$row->{row}] = $rec->{offset};
#      $self->{lengths}->{rows}->[$row->{row}] = $rec->{length};
#    }
#
#    push(@{$rowIndices}, $row->{row});
    my $rowIndex = $self->recordRowRecord($rec);
    push(@{$rowIndices}, $rowIndex);
  }

  $blockInfo->{rowIndices} = $rowIndices;

  return $blockInfo;
}

# Given a block indices structure and a block info structure, marks any rows
#  that were expected within the block but not seen as empty.
sub markMissingRowsInRowBlock($$)
{
  my $self = shift();
  my ($blockIndices, $blockInfo) = @_;

  ($self->{debug}) && printf("markMissingRowsInRowBlock(): blockIndex %d minRowIndex %d maxRowIndex %d\n", $blockIndices->{blockIndex}, $blockIndices->{minRowIndex}, $blockIndices->{maxRowIndex});

  my $seenRows = {};
  for my $rowIndex (@{$blockInfo->{rowIndices}})
  {
    $seenRows->{$rowIndex} = 1;
  }
  
  # Any rows that are within the bounds of this row block but had no Row record are considered empty
#  my $minRow = $self->{meta}->{index}->{minRow} // $index->{minRow};
#  my $minRowIndex = $minRow + ($blockIndex * Spreadsheet::Nifty::XLS::ROW_BLOCK_SIZE);
#  my $maxRowIndex = $minRowIndex + Spreadsheet::Nifty::XLS::ROW_BLOCK_SIZE - 1;
#  ($self->{debug}) && printf("markMissingRowsInRowBlock(): minRow: %d minRowIndex: %d maxRowIndex: %d\n", $minRow, $minRowIndex, $maxRowIndex);
  for (my $i = $blockIndices->{minRowIndex}; $i <= $blockIndices->{maxRowIndex}; $i++)
  {
    if (!defined($seenRows->{$i}))
    {
      ($self->{debug}) && printf("markMissingRowsInRowBlock(): Marking row %d not seen within block as missing\n", $i);
      $self->{offsets}->{rows}->[$i] = -1;
    }
  }

  return;
}

# Given a row index, finds its offset within the file.
# Returns undef if no such row exists in the file.
sub findRow($)
{
  my $self = shift();
  my ($rowIndex) = @_;

  ($self->{debug}) && printf("findRow(): rowIndex %d\n", $rowIndex);

  # Row bounds check
  ($rowIndex < 0) && return undef;
  if (defined($self->{header}->{dimensions}))
  {
    ($rowIndex >= $self->{header}->{dimensions}->{maxRow}) && return undef;  # NOTE: maxRow is actually the first empty trailing row
  }

  # If we already know the offset, use it
  my $offset = $self->{offsets}->{rows}->[$rowIndex];
  if (defined($offset))
  {
    ($self->{debug}) && printf("findRow(): Row offset is known: %d\n", $offset);
    return ($offset >= 0) ? $offset : undef;
  }

  # If we have an Index record, try using the index
  if (defined($self->{header}->{index}))
  {
    ($self->{debug}) && printf("findRow(): Attempting to find via index...\n");
    $offset = $self->findRowViaIndex($rowIndex);
    (defined($offset)) && return ($offset >= 0) ? $offset : undef;
  }

  # Fall back to scanning for the row
  ($self->{debug}) && printf("findRow(): Attempting to find via scan...\n");
  $offset = $self->findRowViaScan($rowIndex);
  return ($offset >= 0) ? $offset : undef;
}

# Given a row index, finds its offset within the file using the Index record.
# Returns undef if the index was not usable.
# Returns a negative value if the row does not exist.
sub findRowViaIndex($)
{
  my $self = shift();
  my ($rowIndex) = @_;

  ($self->{debug}) && printf("findRowViaIndex: %d\n", $rowIndex);

  # Don't bother if there was no index
  my $index = $self->{header}->{index};
  (!defined($index)) && return undef;  # No Index record

  # Row bounds check
  ($rowIndex < $index->{minRow}) && return -1;
  ($rowIndex >= $index->{maxRow}) && return -1;  # NOTE: maxRow is really the first unused row

  my $indices = $self->getRowBlockIndices($rowIndex);
  (!defined($indices)) && return undef;  # Couldn't get row block indices

  my $blockInfo = $self->readRowBlockInfo($indices->{blockIndex});
  (!defined($blockInfo)) && return undef;  # Couldn't get row block info
  ($self->{debug}) && printf("Block info for row %d: %s\n", $rowIndex, main::Dumper($blockInfo));

  $self->markMissingRowsInRowBlock($indices, $blockInfo);

  # If readRowBlockInfo succeeded, it should have filled in the row offsets
  #  for us. So if the offset is now known, just return it.
  my $offset = $self->{offsets}->{rows}->[$rowIndex];
  if (defined($offset))
  {
    return $offset;
  }

  return undef;  # Something went wrong
}

# Given a row index, finds its offset within the file.
# Returns a negative value if the Row record does not exist.
sub findRowViaScan($)
{
  my $self = shift();
  my ($rowIndex) = @_;

  # Try to find an earlier row with a known offset
  my $offset;
  for (my $r = $rowIndex - 1; $r >= 0; $r--)
  {
    if (defined($self->{offsets}->{rows}->[$r]) && ($self->{offsets}->{rows}->[$r] > 0))
    {
      $offset = $self->{offsets}->{rows}->[$r];
      last;
    }
  }

  # If no known offset, start reading at the beginning of the sheet body
  if (!defined($offset))
  {
    $offset = $self->{offsets}->{body};
  }

  # Scan for Row records
  ($self->{debug}) && printf("findRowViaScan(): Looking for row %d starting at offset %d\n", $rowIndex, $offset);
  my $prevRowIndex = undef;
  $self->{biff}->seek($offset);
  while (defined(my $rec = $self->{biff}->readRecord()))
  {
    ($self->{debug}) && printf("findRowViaScan(): Read offset %d rec %s...\n", $rec->{offset}, $rec->{name} // '?');
    ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::EOF) && last;  # End of sheet

    # NOTE: There are files where there is no Row record for a row that does
    #  actually exist. Apparently some writers only write Row records for rows
    #  that have non-default formatting.
    # This could have us scan for a very long time without finding a record.
    #  To work around this, we also decode cell records and stop when we see
    #  one from the target row or a following row.
    # This assumes that a Row record will always precede the cell records for
    #  that same row.
    if ($self->isRecordTypeCell($rec->{type}))
    {
      my $cell = $self->decodeCell($rec);
      (!defined($cell)) && die("findRowViaScan(): Cell decode failed");
      ($self->{debug}) && printf("  It's a cell from row %d\n", $cell->{row});

      # Opportunistically record firstCell offsets
      if (!defined($self->{offsets}->{firstCells}->[$cell->{row}]))
      {
        $self->{offsets}->{firstCells}->[$cell->{row}] = $rec->{offset};
      }

      if ($cell->{row} >= $rowIndex)
      {
        ($self->{debug}) && printf("Found cell from target or following row, giving up scan.\n");
        $self->{offsets}->{rows}->[$rowIndex] = -1;
        return -1;
      }
    }

    ($rec->{type} != Spreadsheet::Nifty::XLS::RecordTypes::ROW) && next;  # Not a Row record

    my $row = Spreadsheet::Nifty::XLS::Decode::decodeRow($rec->{payload});
    if (!defined($self->{offsets}->{rows}->[$row->{row}]))
    {
      $self->{offsets}->{rows}->[$row->{row}] = $rec->{offset};
      $self->{lengths}->{rows}->[$row->{row}] = $rec->{length};
    }

    if ($row->{row} == $rowIndex)
    {
      return $rec->{offset};  # Found target row
    }
    elsif ($row->{row} > $rowIndex)
    {
      # It's a following row, so target row must not exist
      if (defined($prevRowIndex))
      {
        # We can mark every row between the last seen and the new one as non-existent
        for (my $i = $prevRowIndex + 1; $i < $row->{row}; $i++)
        {
          ($self->{debug}) && printf("findRowViaScan(): Marking row %d as missing\n", $i);
          $self->{offsets}->{rows}->[$i] = -1;
        }
      }

      ($self->{debug}) && printf("findRowViaScan(): Found following row %d rather than target row %d\n", $row->{row}, $rowIndex);
      return -1;
    }

    ($self->{debug}) && printf("findRowViaScan(): Found preceding row %d...\n", $row->{row});
    $prevRowIndex = $row->{row};
  }

  return -1;  # No such row in this sheet
}

# Given a row index, finds the offset of the first cell record for that row.
# Returns undef if no such row exists in the file.
sub findRowFirstCell($)
{
  my $self = shift();
  my ($rowIndex) = @_;

  ($self->{debug}) && printf("findRowFirstCell(): rowIndex %d\n", $rowIndex);

  # Bounds check the row index
  ($rowIndex < 0) && return undef;
  if (defined($self->{header}->{dimensions}))
  {
    if ($rowIndex >= $self->{header}->{dimensions}->{maxRow})  # NOTE: maxRow is actually the first empty trailing row
    {
      ($self->{debug}) && printf("findRowFirstCell(): rowIndex >= maxRow %d\n", $self->{header}->{dimensions}->{maxRow});
      return undef;
    }
  }

  # If we already know the offset, use it
  my $offset = $self->{offsets}->{firstCells}->[$rowIndex];
  if (defined($offset))
  {
    ($self->{debug}) && printf("findRowFirstCell(): First cell location is known: %d\n", $offset);
    return ($offset >= 0) ? $offset : undef;
  }

  # If we have an Index record, try using the index
  if (defined($self->{header}->{index}))
  {
    ($self->{debug}) && printf("findRowFirstCell(): Attempting to find via index...\n");
    $offset = $self->findRowFirstCellViaIndex($rowIndex);
    (defined($offset)) && return ($offset >= 0) ? $offset : undef;
  }

  # Fall back to scanning for the first cell
  ($self->{debug}) && printf("findRowFirstCell(): Attempting to find via scan...\n");
  $offset = $self->findRowFirstCellViaScan($rowIndex);
  if (defined($offset))
  {
    ($self->{debug}) && printf("findRowFirstCell(): ...scan found offset %d\n", $offset);
    return ($offset >= 0) ? $offset : undef;
  }

  ($self->{debug}) && printf("findRowFirstCell(): No row found.\n");
  return undef;  # Not found
}

sub findRowFirstCellViaIndex($)
{
  my $self = shift();
  my ($rowIndex) = @_;

  my $indices = $self->getRowBlockIndices($rowIndex);
  (!defined($indices)) && return undef;  # Couldn't get row block indices
  ($self->{debug}) && printf("findRowFirstCellViaIndex(): rowIndex %d block indices: blockIndex %d expected rows within %d..%d\n", $rowIndex, $indices->{blockIndex}, $indices->{minRowIndex}, $indices->{maxRowIndex});

  my $blockInfo = $self->readRowBlockInfo($indices->{blockIndex});
  (!defined($blockInfo)) && return undef;  # Couldn't get row block info
  ($self->{debug}) && printf("Block info for row %d: %s\n", $rowIndex, main::Dumper($blockInfo));

  $self->markMissingRowsInRowBlock($indices, $blockInfo);

  # The dbcell's firstCellOffsets field is an array of relative offsets. The
  #  first element is offset from the end of the first Row record within
  #  the block.
  my $firstRowIndex = $blockInfo->{rowIndices}->[0];
  (!defined($firstRowIndex)) && return undef;
  my $firstRowOffset = $self->{offsets}->{rows}->[$firstRowIndex];
  (!defined($firstRowOffset)) && return undef;
  my $firstRowLength = $self->{lengths}->{rows}->[$rowIndex];
  (!defined($firstRowLength)) && return undef;
  my $offset = $firstRowOffset + 4 + $firstRowLength;  # The firstRowLength is the payload length, so we also add 4 for the BIFF header
  ($self->{debug}) && printf("findRowFirstCellViaIndex(): firstRowIndex %d firstRowOffset %d firstRowLength %d offset %d\n", $firstRowIndex, $firstRowOffset, $firstRowLength, $offset);

  # Find the ordinal position of the target row within this block
  my $position;
  for (my $i = 0; $i < $blockInfo->{rowIndices}; $i++)
  {
    if ($blockInfo->{rowIndices}->[$i] == $rowIndex)
    {
      $position = $i;
      last;
    }
  }

  # If the target row wasn't seen, it must not exist
  (!defined($position)) && return -1;

  # It's possible that we got more rowIndices than firstCellOffsets. In that
  #  case, a Row record exists but the dbcell has no first cell offset for the 
  #  row. This could happen when a row has style information, but no cells.
  if ($position >= scalar(@{$blockInfo->{dbcell}->{firstCellOffsets}}))
  {
    ($self->{debug}) && printf("findRowFirstCellViaIndex(): position %d but dbcell only had %d entries, assuming no cells within row\n", $position, scalar(@{$blockInfo->{dbcell}->{firstCellOffsets}}));
    $self->{offsets}->{firstCells}->[$rowIndex] = -1;
    return -1;
  }

  # Find the absolute offset of the target row
  for (my $i = 0; $i <= $position; $i++)
  {
    # Relative offsets should not be zero except perhaps the first one.
    # OpenOffice says: If the size of all cell records of a row exceeds
    #  FFFF H, the respective position in the DBCELL record will contain the
    #  offset 0000H. From this point on, the offsets cannot be used anymore
    #  to calculate stream positions.
    (($i > 0) && ($blockInfo->{dbcell}->{firstCellOffsets}->[$i] == 0)) && return undef;  # Index not usable for this row

    $offset = $blockInfo->{dbcell}->{firstCellOffsets}->[$i] + $offset;
  }

  # Read the record at the offset
  $self->{biff}->seek($offset);
  my $rec = $self->{biff}->readRecord();
  (!defined($rec)) && return undef;  # Couldn't read record
  
  # Skip over any leading 'Uncalced' cell
  if (($rec->{name} // '') eq 'Uncalced')
  {
    $rec = $self->{biff}->readRecord();
    (!defined($rec)) && return undef;  # Couldn't read record
  }

  # Make sure it's a cell record from the expected row
  my $cell = $self->decodeCell($rec);
  (!defined($cell)) && return undef;  # Not a cell
  ($cell->{row} != $rowIndex) && return undef;  # Wrong row

  # Record the offset
  $self->{offsets}->{firstCells}->[$rowIndex] = $rec->{offset};

  return $offset;
}

sub findRowFirstCellViaScan($)
{
  my $self = shift();
  my ($rowIndex) = @_;

  # Try to find the offset of the Row record for this row
  my $offset = $self->findRow($rowIndex);
  if (!defined($offset))
  {
    # NOTE: Some writers do not include Row records for all rows with data.
    #  The findRow() scan can end upon finding a cell within the target (or
    #  following) row, rather than a Row record. So before we fail due to
    #  the lack of a Row record, we check one last time that the first cell
    #  offset is not yet known.
    (defined($self->{offsets}->{firstCells}->[$rowIndex])) && return $self->{offsets}->{firstCells}->[$rowIndex];
    return -1;  # No such Row record in file
  }
  
  # Read the Row record and make sure it's the correct one
  $self->{biff}->seek($offset);
  my $rec = $self->{biff}->readRecord();
  (!defined($rec)) && die("findRowFirstCellViaScan(): Expected a Row record, got EOF");
  (($rec->{name} // '') ne 'Row') && die(sprintf("findRowFirstCellViaScan(): Expected a Row record, got 0x%04X (%s)", $rec->{type}, $rec->{name}));
  my $row = Spreadsheet::Nifty::XLS::Decode::decodeRow($rec->{payload});
  ($rowIndex != $row->{row}) && die(sprintf("Expected Row record for row %d, got row %d", $rowIndex, $row->{row}));

  # Skip past the following Row records in this block, if any
  while (defined($rec = $self->{biff}->readRecord()))
  {
    #($rec->{name} ne 'Row') && last;  # Not a Row record
    ($rec->{type} != Spreadsheet::Nifty::XLS::RecordTypes::ROW) && last;  # Not a Row record
    my $skipRow = Spreadsheet::Nifty::XLS::Decode::decodeRow($rec->{payload});
    ($self->{debug}) && printf("findRowFirstCellViaScan(): Skipping Row record %d...\n", $skipRow->{row});

    # Not strictly necessary, but we might as well opportunistically record the Row record positions
    if (!defined($self->{offsets}->{rows}->[$skipRow->{row}]))
    {
      $self->{offsets}->{rows}->[$skipRow->{row}] = $rec->{offset};
      $self->{lengths}->{rows}->[$skipRow->{row}] = $rec->{length};
    }

    # NOTE: Normally we see Row records in blocks of 32, but there are files
    #  with tens of thousands of adjacent Row records covering empty cells, so
    #  we don't want to read Row records forever. We need to allow for some
    #  Row records beyond the target row, because they appear in blocks, but
    #  we also want to stop at a reasonable point.
    if ($skipRow->{row} > ($rowIndex + (Spreadsheet::Nifty::XLS::ROW_BLOCK_SIZE * 4)))
    {
      # There must not be any cells on the target row, but also there must
      #  not be any on quite a few following rows! We'll be conservative here
      #  and only mark enough rows for a single block as missing.
      # In theory, since Row records come in blocks of 32, anything >=
      #  $rowIndex and < ($skipRow->{row} - XLS::ROW_BLOCK_SIZE) could be marked.
      for (my $i = $rowIndex; $i < ($rowIndex + Spreadsheet::Nifty::XLS::ROW_BLOCK_SIZE); $i++)
      {
        ($self->{debug}) && printf("findRowFirstCellViaScan(): Marking row %d as missing due to row overshoot\n", $i);
        $self->{offsets}->{firstCells}->[$i] = -1;
      }  
      return -1;
    }

    # I have seen Row records wrap back around to zero after 65535.
    if ($skipRow->{row} < $rowIndex)
    {
      ($self->{debug}) && printf("findRowFirstCellViaScan(): Row wraparound? Ridiculous.\n");
      $self->{offsets}->{firstCells}->[$rowIndex] = -1;
      return -1;
    }
  }

  # Decode cell records, noting the first cell on each row
  my $prevCellRow = undef;
  while (1)
  {
    (!defined($rec)) && return undef;  # No more records, never found target row
    #(($rec->{name} // '') eq 'Row') && return undef;  # Hit a new block of rows
    ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::ROW) && return undef;  # Hit a new block of rows
    #(($rec->{name} // '') eq 'EOF') && return undef;  # Hit end of of sheet
    ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::EOF) && return undef;  # Hit end of sheet

    # If it's not a cell record, it's not interesting
    my $cell = $self->decodeCell($rec);
    if (!defined($cell))
    {
      ($self->{debug}) && printf("findRowFirstCellViaScan(): Skipping non-cell record with type 0x%X...\n", $rec->{type});
      $rec = $self->{biff}->readRecord();
      next;   
    }

    ($self->{debug}) && printf("findRowFirstCellViaScan(): Found cell type 0x%X row %d\n", $cell->{type}, $cell->{row});

    # If it's from the same row as the previous cell, it's not interesting
    if (defined($prevCellRow) && ($prevCellRow == $cell->{row}))
    {
      $rec = $self->{biff}->readRecord();
      next;   
    }

    # Any gaps between the previous cell's row and this cell's row must be missing rows
    if (defined($prevCellRow))
    {
      for (my $i = $prevCellRow + 1; $i < $cell->{row}; $i++)
      {
        ($self->{debug}) && printf("findRowFirstCellViaScan(): Marking row %d as missing due to cell overshoot\n", $i);
        $self->{offsets}->{firstCells}->[$i] = -1;
      }
    }

    # Record this row's first cell
    $self->{offsets}->{firstCells}->[$cell->{row}] = $rec->{offset};
    $prevCellRow = $cell->{row};

    if ($cell->{row} == $rowIndex)
    {
      return $rec->{offset};  # Found the target row
    }
    elsif ($cell->{row} > $rowIndex)
    {
      return -1;  # Past the target row, so the target must not exist
    }
  }
}

# Reads cell records for the given row.
sub readRowCells($)
{
  my $self = shift();
  my ($rowIndex) = @_;

  my $offset = $self->findRowFirstCell($rowIndex);
  ($self->{debug}) && printf("readRowCells(): rowIndex %d offset %s\n", $rowIndex, $offset // '(none)');
  (!defined($offset)) && return [];  # Can't find row

  my $cells = [];
  $self->{biff}->seek($offset);
  while (defined(my $rec = $self->{biff}->readRecord()))
  {
    ($self->{debug}) && print main::Dumper('REC', $rec);
    ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::UNCALCED) && next;  # These can appear before cells holding formulas. We ignore them.
    ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::EOF) && last;  # Hit EOF record (end of substream)

    if ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::ROW)
    {
      # Hit a new row block
      $self->recordRowRecord($rec);
      last;
    }

    my $cell = $self->decodeCell($rec);
    (!defined($cell)) && next;  # Not a cell record

    if ($cell->{type} == Spreadsheet::Nifty::XLS::RecordTypes::FORMULA)  # A Formula record can have various extra records following
    {
      # Extra formula data may follow
      my $extra = $self->{biff}->readRecordIfType([Spreadsheet::Nifty::XLS::RecordTypes::SHARED_FORMULA, Spreadsheet::Nifty::XLS::RecordTypes::ARRAY, Spreadsheet::Nifty::XLS::RecordTypes::TABLE]);
      if (defined($extra))
      {
        $cell->{extra} = $extra->{type};
        if ($extra->{type} == Spreadsheet::Nifty::XLS::RecordTypes::SHARED_FORMULA)
        {
          $cell->{shared} = Spreadsheet::Nifty::XLS::Decode::decodeShrFmla($extra->{payload});
          ($self->{debug}) && print main::Dumper('shrFmla', $cell->{shared});

          my $tokens = Spreadsheet::Nifty::XLS::Formula::decodeTokens($cell->{shared}->{formula}, $self->getBiffVersion());
          my $context = {sheet => $self, row => $rowIndex, col => $cell->{col}};
          #printf("Shared formula: %s\n", Spreadsheet::Nifty::XLS::Formula::unparseTokens($context, $tokens));
          #($self->{debug}) && Spreadsheet::Nifty::XLS::Formula::decodeTokens($cell->{shared}->{formula}, $self->getBiffVersion());
        }

        ($self->{debug}) && print main::Dumper('EXTRA', $extra);
      }

      # If the result of the formula is a string, it is followed by a String record
      my $string = $self->{biff}->readRecordIfType(Spreadsheet::Nifty::XLS::RecordTypes::STRING);
      if (defined($string))
      {
        ($self->{debug}) && print main::Dumper('STRING', $string);
        $cell->{string} = Spreadsheet::Nifty::XLS::Decode::decodeString($string->{payload}, $self->getBiffVersion());
      }
    }

    ($self->{debug}) && print main::Dumper('CELL', $cell);

    if ($cell->{row} > $rowIndex)
    {
      # We've reached a row beyond the target row
      $self->{offsets}->{firstCells}->[$cell->{row}] = $rec->{offset};
      last;
    }
    elsif ($cell->{row} == $rowIndex)
    {
      push(@{$cells}, $cell);
    }
    else
    {
      die("readRowCells(): Expected rows >= $rowIndex, got $cell->{row}");
    }
  }

  return $cells;
}

# Given a raw Row record (not yet decoded), records row details if not already known.
# Returns the row index from the record.
sub recordRowRecord($)
{
  my $self = shift();
  my ($rec) = @_;

  ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::ROW) || die("recordRowRecord(): Given a non-Row record");

  # We don't bother decoding the whole thing because we just care about the rowIndex
  my $rowIndex = unpack('v', $rec->{payload});

  if (!defined($self->{offsets}->{rows}->[$rowIndex]))
  {
    ($self->{debug}) && printf("recordRowRecord(): Recording record for rowIndex %d\n", $rowIndex);
    $self->{offsets}->{rows}->[$rowIndex] = $rec->{offset};
    $self->{lengths}->{rows}->[$rowIndex] = $rec->{length};
  }

  return $rowIndex;
}

# Given a record type, returns true if it is a cell record.
sub isRecordTypeCell($)
{
  my $self = shift();
  my ($type) = @_;

  CORE::state $cellTypes =
  {
    Spreadsheet::Nifty::XLS::RecordTypes::BLANK     => 1,
    Spreadsheet::Nifty::XLS::RecordTypes::LABEL     => 1,
    Spreadsheet::Nifty::XLS::RecordTypes::LABEL_SST => 1,
    Spreadsheet::Nifty::XLS::RecordTypes::RSTRING   => 1,
    Spreadsheet::Nifty::XLS::RecordTypes::RK        => 1,
    Spreadsheet::Nifty::XLS::RecordTypes::NUMBER    => 1,
    Spreadsheet::Nifty::XLS::RecordTypes::BOOL_ERR  => 1,
    Spreadsheet::Nifty::XLS::RecordTypes::MUL_BLANK => 1,
  };

  return defined($cellTypes->{$type});
}

# Given a BIFF record for a cell, returns a cell strcture.
# Returns undef if the record was not a cell record.
sub decodeCell($)
{
  my $self = shift();
  my ($rec) = @_;

  my $cell;
  if ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::BLANK)
  {
    $cell = Spreadsheet::Nifty::XLS::Decode::decodeBlank($rec->{payload});
  }
  elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::LABEL)
  {
    $cell = Spreadsheet::Nifty::XLS::Decode::decodeLabel($rec->{payload}, $self->getBiffVersion());
  }
  elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::LABEL_SST)
  {
    $cell = Spreadsheet::Nifty::XLS::Decode::decodeLabelSst($rec->{payload});
  }
  elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::RSTRING)
  {
    $cell = Spreadsheet::Nifty::XLS::Decode::decodeRString($rec->{payload}, $self->getBiffVersion());
  }
  elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::RK)
  {
    $cell = Spreadsheet::Nifty::XLS::Decode::decodeRK($rec->{payload});
  }
  elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::NUMBER)
  {
    $cell = Spreadsheet::Nifty::XLS::Decode::decodeNumber($rec->{payload});
  }
  elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::BOOL_ERR)
  {
    $cell = Spreadsheet::Nifty::XLS::Decode::decodeBoolErr($rec->{payload});
  }
  elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::MUL_BLANK)
  {
    $cell = Spreadsheet::Nifty::XLS::Decode::decodeMulBlank($rec->{payload});
  }
  elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::MUL_RK)
  {
    $cell = Spreadsheet::Nifty::XLS::Decode::decodeMulRk($rec->{payload});
  }
  elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::FORMULA)
  {
    $cell = Spreadsheet::Nifty::XLS::Decode::decodeFormula($rec->{payload});
  }
  else
  {
    return undef;
  }

  $cell->{type} = $rec->{type};
  return $cell;
}

sub getRowDimensions()
{
  my $self = shift();

  return [$self->{header}->{dimensions}->{minRow}, $self->{header}->{dimensions}->{maxRow} - 1];
}

sub getColDimensions()
{
  my $self = shift();

  return [$self->{header}->{dimensions}->{minCol}, $self->{header}->{dimensions}->{maxCol} - 1];
}

1;
