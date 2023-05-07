#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSB::SheetReader;

# === Class methods ===

sub new()
{
  my $class = shift();
  my ($workbook, $index, $partname) = @_;

  my $self = {};
  $self->{workbook}   = $workbook;
  $self->{sheetIndex} = $index;
  $self->{partname}   = $partname;
  $self->{debug}      = $workbook->{debug};
  $self->{records}    = undef;
  $self->{rowIndex}   = 0;
  $self->{rowHeader}  = undef;  # BrtRowHdr or BrtEndSheetData
  $self->{protection} = undef;  # BrtSheetProtection or BrtSheetProtectionIso
  $self->{props}      = {};  # BrtWsProp

  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub open()
{
  my $self = shift();

  my $reader = $self->{workbook}->{zipPackage}->openMember($self->{partname});
  (!$reader) && die("Couldn't open member '$self->{partname}'\n");

  $self->{records} = Spreadsheet::Nifty::XLSB::RecordReader->new($reader);

  # Read records up until 'BrtBeginSheetData'
  while (my $rec = $self->{records}->read())
  {
    ($self->{debug}) && printf("SheetReader open(): REC type %d size %d name %s\n", $rec->{type}, $rec->{size}, $rec->{name} // '?');
    ($self->{debug}) && ($rec->{size}) && printf("  data: %s\n", unpack('H*', $rec->{data}));

    if ($rec->{name} eq 'BrtWsProp')
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $wsprop = $decoder->decodeHash(['flags:u24', 'brtcolorTab:bytes[8]', 'rwSync:bytes[4]', 'colSync:bytes[4]', 'strName:XLWideString']);
      $self->{props} = $wsprop;
      ($self->{debug}) && print main::Dumper('wsprop', $wsprop);
    }
    elsif ($rec->{name} eq 'BrtWsDim')
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $wsdim = $decoder->decodeHash(['minrow:u32', 'maxrow:u32', 'mincol:u32', 'maxcol:u32']);
      $self->{dimensions} = $wsdim;
      ($self->{debug}) && print main::Dumper('wsdim', $wsdim);
    }
    elsif ($rec->{name} eq 'BrtBeginSheetData')
    {
      $self->readRowHeader();
      return;
    }
  }

  return;
}

# NOTE: Resets the read position. Don't attempt to read rows after calling. Don't attempt to call twice.
sub readProtection()
{
  my $self = shift();

  # TODO: Use binary index or cached offsets

  # Skip over any remaining row data
  while (defined($self->readRow())) {}

  # Read records looking for 'BrtSheetProtectionIso' or 'BrtSheetProtection'
  while (my $rec = $self->{records}->read())
  {
    if ($rec->{name} eq 'BrtSheetProtectionIso')
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $protection = $decoder->decodeHash(['spinCount:u32', ':bytes[64]']);  # TODO: Flag fields
      my $hashLength = $decoder->decodeField('u32');
      $protection->{hash} = $decoder->getBytes($hashLength);
      my $saltLength = $decoder->decodeField('u32');
      $protection->{salt} = $decoder->getBytes($saltLength);
      $protection->{type} = $decoder->decodeField('XLNullableWideString');

      print main::Dumper($protection);
      $self->{protection} = {type => $rec->{name}, data => $protection};
      last;
    }
    elsif ($rec->{name} eq 'BrtSheetProtection')
    {
      (defined($self->{protection})) && next;  # Likely already has BrtSheetProtectionIso

      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $protection = $decoder->decodeHash(['hash:u16']);  # TODO: Flag fields
      $protection->{type} = 'excel16';

      $self->{protection} = {type => $rec->{name}, data => $protection};
      last;
    }
  }

  return $self->{protection}->{data};
}

# Reads records up until BrtRowHdr or BrtEndSheetData
sub readRowHeader()
{
  my $self = shift();

  my $rowHdr;
  my $row = [];

  while (my $rec = $self->{records}->read())
  {
    ($self->{debug}) && printf("SheetReader readRowHeader(): REC type %d size %d name %s\n", $rec->{type}, $rec->{size}, $rec->{name} // '?');
    ($self->{debug}) && ($rec->{size}) && printf("  data: %s\n", unpack('H*', $rec->{data}));

    if ($rec->{name} eq 'BrtRowHdr')
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $rowHdr = $decoder->decodeHash(['row:u32', 'ixfe:u32', 'height:u16', 'flags1:u16', 'flags2:u8', 'ccolspan:u32']);
      $self->{rowHeader} = {type => $rec->{name}, data => $rowHdr};
      ($self->{debug}) && print main::Dumper($rowHdr);
      last;
    }
    elsif ($rec->{name} eq 'BrtEndSheetData')
    {
      $self->{rowHeader} = {type => $rec->{name}};
      last;
    }
  }

  return;
}

# Given a cell record, returns a cell structure, or undef.
sub decodeCell($)
{
  my $self = shift();
  my ($rec) = @_;

  my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
  my $cell;

  if ($rec->{name} eq 'BrtCellBlank')
  {
    $cell = $decoder->decodeHash(['column:u32', 'iStyleRef:u24', 'flags:u8']);
    $cell->{dataType} = Spreadsheet::Nifty::TYPE_NULL;
  }
  elsif ($rec->{name} eq 'BrtCellBool')
  {
    $cell = $decoder->decodeHash(['column:u32', 'iStyleRef:u24', 'flags:u8', 'value:u8']);
    $cell->{dataType} = Spreadsheet::Nifty::TYPE_BOOL;
  }
  elsif ($rec->{name} eq 'BrtCellError')
  {
    $cell = $decoder->decodeHash(['column:u32', 'iStyleRef:u24', 'flags:u8', 'value:u8']);
    $cell->{dataType} = Spreadsheet::Nifty::TYPE_ERR;
  }
  elsif ($rec->{name} eq 'BrtCellSt')
  {
    $cell = $decoder->decodeHash(['column:u32', 'iStyleRef:u24', 'flags:u8', 'value:XLWideString']);
    $cell->{dataType} = Spreadsheet::Nifty::TYPE_STR;
  }
  elsif ($rec->{name} eq 'BrtCellRString')  # Rich string
  {
    $cell = $decoder->decodeHash(['column:u32', 'iStyleRef:u24', 'flags:u8', 'flags2:u8', 'value:XLWideString']);  # TODO: More fields
    $cell->{dataType} = Spreadsheet::Nifty::TYPE_STR;
  }
  elsif ($rec->{name} eq 'BrtCellReal')
  {
    $cell = $decoder->decodeHash(['column:u32', 'iStyleRef:u24', 'flags:u8', 'value:f64']);
    $cell->{dataType} = Spreadsheet::Nifty::TYPE_NUM;
    #printf("  Real: %g\n", $cell->{value});
  }
  elsif ($rec->{name} eq 'BrtCellIsst')
  {
    $cell = $decoder->decodeHash(['column:u32', 'iStyleRef:u24', 'flags:u8', 'value:u32']);
    my $str = $self->{workbook}->{sharedStrings}->[$cell->{value}]->{string};
    ($self->{debug}) && printf("Interned string: %s\n", $str);
    $cell->{dataType} = Spreadsheet::Nifty::TYPE_STR;
    $cell->{value}    = $str;
  }
  elsif ($rec->{name} eq 'BrtCellRk')
  {
    $cell = $decoder->decodeHash(['column:u32', 'iStyleRef:u24', 'flags:u8']);
    my $value = Spreadsheet::Nifty::Utils->translateRk($decoder->decodeField('u32'));
    $cell->{dataType} = Spreadsheet::Nifty::TYPE_NUM;
    $cell->{value} = $value;
  }
  elsif ($rec->{name} eq 'BrtFmlaBool')
  {
    $cell = $decoder->decodeHash(['column:u32', 'iStyleRef:u24', 'flags:u8', 'value:u8', 'flags2:u16']);  # TODO: More fields
    $cell->{dataType} = Spreadsheet::Nifty::TYPE_BOOL;
  }
  elsif ($rec->{name} eq 'BrtFmlaError')
  {
    $cell = $decoder->decodeHash(['column:u32', 'iStyleRef:u24', 'flags:u8', 'value:u8', 'flags2:u16']);  # TODO: More fields
    $cell->{dataType} = Spreadsheet::Nifty::TYPE_ERR;
  }
  elsif ($rec->{name} eq 'BrtFmlaNum')
  {
    $cell = $decoder->decodeHash(['column:u32', 'iStyleRef:u24', 'flags:u8', 'value:f64']);  # TODO: More fields
    $cell->{dataType} = Spreadsheet::Nifty::TYPE_NUM;
  }
  elsif ($rec->{name} eq 'BrtFmlaString')
  {
    $cell = $decoder->decodeHash(['column:u32', 'iStyleRef:u24', 'flags:u8', 'value:XLWideString', 'flags2:u16']);  # TODO: More fields
    $cell->{dataType} = Spreadsheet::Nifty::TYPE_STR;
  }

  (!defined($cell)) && return undef;

  $cell->{type} = $rec->{type};
  
  return $cell;
}

# Reads a row of cells
sub readRowInternal()
{
  my $self = shift();

  my $row = [];

  while (my $rec = $self->{records}->read())
  {
    ($self->{debug}) && printf("SheetReader readRowInternal(): REC type %d size %d name %s\n", $rec->{type}, $rec->{size}, $rec->{name} // '?');
    ($self->{debug}) && ($rec->{size}) && printf("  data: %s\n", unpack('H*', $rec->{data}));

    if ($rec->{name} eq 'BrtRowHdr')
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $rowHdr = $decoder->decodeHash(['row:u32', 'ixfe:u32', 'height:u16', 'flags1:u16', 'flags2:u8', 'ccolspan:u32']);
      $self->{rowHeader} = {type => $rec->{name}, data => $rowHdr};
      ($self->{debug}) && print main::Dumper($rowHdr);
      last;
    }
    elsif ($rec->{name} eq 'BrtEndSheetData')
    {
      $self->{rowHeader} = {type => $rec->{name}};
      last;
    }
    else
    {
      my $cell = $self->decodeCell($rec);
      (!defined($cell)) && next;

      (!defined($cell->{column})) && die("Returned cell with no column?");
      $row->[$cell->{column}] = $cell;
    }
  }

  return $row;
}

sub tellRow()
{
  my $self = shift();

  return $self->{rowIndex};
}

sub readRow()
{
  my $self = shift();

  (!defined($self->{rowHeader})) && die("No row header?");

  ($self->{rowHeader}->{type} eq 'BrtEndSheetData') && return undef;  # No more rows

  ($self->{rowIndex} > $self->{rowHeader}->{data}->{row}) && die("Unordered row headers?");

  if ($self->{rowIndex} < $self->{rowHeader}->{data}->{row})
  {
    # We haven't yet reached the next row containing data
    $self->{rowIndex}++;
    return [];
  }

  my $row = $self->readRowInternal();
  (!defined($row)) && return undef;  # Out of rows

  $self->{rowIndex}++;
  return $row;
}

sub readRowData()
{
  my $self = shift();

  my $row = $self->readRow();
  (!defined($row)) && return undef;  # Out of rows

  $row = [ map({ defined($_) ? $_->{v} : undef } @{$row}) ];  # Throw away everything but values

  return $row;
}

sub getName()
{
  my $self = shift();

  return $self->{workbook}->getSheetNames()->[$self->{sheetIndex}];
}

sub getRowDimensions()
{
  my $self = shift();

  return [$self->{dimensions}->{minrow}, $self->{dimensions}->{maxrow}];
}

sub getColDimensions()
{
  my $self = shift();

  return [$self->{dimensions}->{mincol}, $self->{dimensions}->{maxcol}];
}

1;
