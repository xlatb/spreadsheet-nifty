#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::BIFF::FormulaDecoder;

use Spreadsheet::Nifty::BIFF::Formula;
use Spreadsheet::Nifty::BIFF::FormulaTokens;
use Spreadsheet::Nifty::BIFF::Constants;
use Spreadsheet::Nifty::BIFF::Functions;

use Encode qw();

# === Class methods ===

sub new($)
{
  my $class = shift();
  my ($version) = @_;

  my $self = {};
  $self->{version} = $version;

  if ($version == Spreadsheet::Nifty::BIFF::Constants::BIFF_V5)
  {
    $self->{maxColCount} = 256;    # 8 bits
    $self->{maxRowCount} = 16384;  # 14 bits
  }
  elsif ($version == Spreadsheet::Nifty::BIFF::Constants::BIFF_V8)
  {
    $self->{maxColCount} = 256;    # 8 bits
    $self->{maxRowCount} = 65536;  # 16 bits
  }
  elsif ($version == Spreadsheet::Nifty::BIFF::Constants::BIFF_V12)
  {
    $self->{maxColCount} = 16384;    # 14 bits
    $self->{maxRowCount} = 1048576;  # 20 bits
  }
  else
  {
    die("Unknown BIFF version");
  }

  bless($self, $class);
  return $self;
}

# === Instance methods ===

sub decodeCellRef2d($)
{
  my $self = shift();
  my ($decoder) = @_;

  my $ref = {};

  if ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V5)
  {
    my $row = $decoder->decodeField('u16');
    my $col = $decoder->decodeField('u8');
    $ref->{row} = $row & 0x3FFF;
    $ref->{col} = $col;
    $ref->{relRow} = ($row >> 15) & 0x1;
    $ref->{relCol} = ($row >> 14) & 0x1;
  }
  elsif ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V8)
  {
    my $row = $decoder->decodeField('u16');
    my $col = $decoder->decodeField('u16');
    $ref->{row} = $row;
    $ref->{col} = $col & 0xFF;
    $ref->{relRow} = ($col >> 15) & 0x1;
    $ref->{relCol} = ($col >> 14) & 0x1;
  }
  elsif ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V12)
  {
    my $row = $decoder->decodeField('u32');
    my $col = $decoder->decodeField('u16');
    $ref->{row} = $row;
    $ref->{col} = $col & 0x3FFF;
    $ref->{relRow} = ($col >> 15) & 0x1;
    $ref->{relCol} = ($col >> 14) & 0x1;
  }
  else
  {
    ...;
  }

  return $ref;
}

sub decodeCellRef3d($)
{
  my $self = shift();
  my ($decoder) = @_;

  my $ref;

  if ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V5)
  {
    ...;  # TODO
  }
  elsif ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V8)
  {
    my $sheetIndex = $decoder->decodeField('u16');  # Index into EXTERNSHEET
    $ref = $self->decodeCellRef2d($decoder);
    $ref->{sheetIndex} = $sheetIndex;
  }
  elsif ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V12)
  {
    my $sheetIndex = $decoder->decodeField('u16');
    $ref = $self->decodeCellRef2d($decoder);
    $ref->{sheetIndex} = $sheetIndex;
  }
  else
  {
    ...;
  }

  return $ref;
}

sub decodeAreaRef2d($)
{
  my $self = shift();
  my ($decoder) = @_;

  my $ref = {};

  if ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V5)
  {
    my $minRow = $decoder->decodeField('u16');
    my $maxRow = $decoder->decodeField('u16');
    my $minCol = $decoder->decodeField('u8');
    my $maxCol = $decoder->decodeField('u8');
    $ref->{minRow} = $minRow & 0x3FFF;
    $ref->{maxRow} = $maxRow & 0x3FFF;
    $ref->{minCol} = $minCol;
    $ref->{maxCol} = $maxCol;
    $ref->{relMinRow} = ($minRow >> 15) & 0x1;
    $ref->{relMinCol} = ($minRow >> 14) & 0x1;
    $ref->{relMaxRow} = ($maxRow >> 15) & 0x1;
    $ref->{relMaxCol} = ($maxRow >> 14) & 0x1;
  }
  elsif ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V8)
  {
    my $minRow = $decoder->decodeField('u16');
    my $maxRow = $decoder->decodeField('u16');
    my $minCol = $decoder->decodeField('u16');
    my $maxCol = $decoder->decodeField('u16');
    $ref->{minRow} = $minRow;
    $ref->{maxRow} = $maxRow;
    $ref->{minCol} = $minCol & 0xFF;
    $ref->{maxCol} = $maxCol & 0xFF;
    $ref->{relMinRow} = ($minCol >> 15) & 0x1;
    $ref->{relMinCol} = ($minCol >> 14) & 0x1;
    $ref->{relMaxRow} = ($maxCol >> 15) & 0x1;
    $ref->{relMaxCol} = ($maxCol >> 14) & 0x1;
  }
  elsif ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V12)
  {
    my $minRow = $decoder->decodeField('u32');
    my $maxRow = $decoder->decodeField('u32');
    my $minCol = $decoder->decodeField('u16');
    my $maxCol = $decoder->decodeField('u16');
    $ref->{minRow} = $minRow;
    $ref->{maxRow} = $maxRow;
    $ref->{minCol} = $minCol & 0x3FFF;
    $ref->{maxCol} = $maxCol & 0x3FFF;
    $ref->{relMinRow} = ($minCol >> 15) & 0x1;
    $ref->{relMinCol} = ($minCol >> 14) & 0x1;
    $ref->{relMaxRow} = ($maxCol >> 15) & 0x1;
    $ref->{relMaxCol} = ($maxCol >> 14) & 0x1;
  }
  else
  {
    ...;
  }

  return $ref;
}

sub decode($$)
{
  my $self = shift();
  my ($tokenBytes, $extraBytes) = @_;

  my $decoder = StructDecoder->new($tokenBytes);
  my $extraDecoder = StructDecoder->new($extraBytes // '');

  my $formula = Spreadsheet::Nifty::BIFF::Formula->new($self->{maxColCount}, $self->{maxRowCount});

  while ($decoder->bytesLeft())
  {
    my $type = $decoder->decodeField('u8');
    my $class = ($type & Spreadsheet::Nifty::BIFF::FormulaTokens::TOKEN_MASK_CLASS) >> Spreadsheet::Nifty::BIFF::FormulaTokens::TOKEN_SHIFT_CLASS;
    my $id    = ($type & Spreadsheet::Nifty::BIFF::FormulaTokens::TOKEN_MASK_ID);

    my $token = {type => $type, class => $class, id => $id};

    my $info = Spreadsheet::Nifty::BIFF::FormulaTokens::tokenInfo($type);
    (!defined($info)) && die(sprintf("Unknown token type 0x%02X class %d id 0x%02X", $type, $class, $id));

    #printf("FormulaDecoder->decode(): Token 0x%02X class %d id 0x%02X name '%s'\n", $type, $class, $id, $info->{name});

    if ($class == Spreadsheet::Nifty::BIFF::FormulaTokens::TOKEN_CLASS_BASIC)  # Basic tokens
    {
      if ($id == 0x01)
      {
        # tExp - References a shared formula on the current sheet
        # If it appears, it must be the first token within a formula.
        if ($self->{version} <= Spreadsheet::Nifty::BIFF::Constants::BIFF_V8)
        {
          $token->{exp} = $decoder->decodeHash(['row:u16', 'col:u16']);
        }
        elsif ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V12)
        {
          # NOTE: BIFF12 splits the row and column numbers between the main stream and extra stream
          $token->{exp}->{row} = $decoder->decodeField('u32');
          $token->{exp}->{col} = $extraDecoder->decodeField('u32');
        }
      }
      elsif (($id >= 0x03) && ($id <= 0x16))
      {
        # Trivial single-byte tokens
      }
      elsif ($id == 0x17)
      {
        # tStr
        if ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V5)
        {
          $token->{str} = Spreadsheet::Nifty::XLS::Decode::decodeLen8AnsiString($decoder);
        }
        elsif ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V8)
        {
          $token->{str} = Spreadsheet::Nifty::XLS::Decode::decodeShortXLUnicodeString($decoder);
        }
        else
        {
          my $size = $decoder->decodeField('u16');
          my $payload = $decoder->getBytes($size * 2);
          $token->{str} = Encode::decode('UTF-16LE', $payload);
        }
      }
      elsif ($id == 0x19)
      {
        # tAttr - Special attributes, meaning depends on flags byte
        my $flags = $decoder->decodeField('u8');
        $token->{attr}->{flags} = $flags;
        #printf("  flags 0x%02X\n", $flags);

        # Many of these can be ignored for our purposes, but we still need to
        #  account for each flags possiblity to know how many bytes to read.
        if ($flags == 0x01)
        {
          # tAttrVolatile
          #printf("  tAttrVolatile\n");
          $decoder->decodeField('u16');  # Unused
        }
        elsif ($flags == 0x02)
        {
          # tAttrIf
          #printf("  tAttrIf\n");
          $token->{attr}->{skipCount} = $decoder->decodeField('u16');
        }
        elsif ($flags == 0x04)
        {
          # tAttrChoose
          #printf("  tAttrChoose\n");
          $token->{attr}->{choose} = $decoder->decodeHash(['choiceCount:u16', 'jumpTable:u16[choiceCount]', 'skipCount:u16']);
        }
        elsif ($flags == 0x08)
        {
          # tAttrSkip
          #printf("  tAttrSkip\n");
          $token->{attr}->{skipCount} = $decoder->decodeField('u16');
        }
        elsif ($flags == 0x10)
        {
          # tAttrSum - SUM() with one parameter
          $decoder->getBytes(2);  # Skip padding
        }
        elsif ($flags == 0x020)
        {
          # tAttrAssign - Assignment within macro sheet?
          $decoder->getBytes(2);  # Skip padding
        }
        elsif (($flags == 0x040) || ($flags == 0x41))
        {
          # tAttrSpace & tAttrSpaceVolatile - Whitespace
          #printf("  tAttrSpace or tAttrSpaceVolatile\n");
          $token->{attr}->{space} = $decoder->decodeHash(['type:u8', 'count:u8']);
        }
        else
        {
          ...;
        }
      }
      elsif ($id == 0x1D)
      {
        # tBool
        $token->{value} = $decoder->decodeField('u8');
      }
      elsif ($id == 0x1E)
      {
        # tInt
        $token->{value} = $decoder->decodeField('u16');
      }
      elsif ($id == 0x1F)
      {
        # tNum
        $token->{value} = $decoder->decodeField('f64');
      }
      else
      {
        die(sprintf("decode(): Basic token 0x%02X unknown", $id));
        ...;
      }
    }
    else  # Classified tokens
    {
      if ($id == 0x01)
      {
        # tFunc
        $token->{func} = $decoder->decodeHash(['func:u16']);
      }
      elsif ($id == 0x02)
      {
        # tFuncVar
        $token->{func} = $decoder->decodeHash(['argCount:u8', 'func:u16']);
      }
      elsif ($id == 0x03)
      {
        # tName
        if ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V5)
        {
          $token->{index} = $decoder->decodeField('u16');
          $decoder->getBytes(12);  # Trailing padding
        }
        elsif ($self->{version} == Spreadsheet::Nifty::BIFF::Constants::BIFF_V8)
        {
          $token->{index} = $decoder->decodeField('u16');
          $decoder->getBytes(2);  # Trailing padding
        }
        else
        {
          ...;
        }
      }
      elsif (($id == 0x04) || ($id == 0x0C))
      {
        # tRef - Reference to another cell
        # tRefN - Reference relative to shared formula base cell
        $token->{ref} = $self->decodeCellRef2d($decoder);
      }
      elsif (($id == 0x05) || ($id == 0x0D))
      {
        # tArea - Reference to a cell range on the same sheet
        # tAreaN - Reference to cell range relative to shared formula base cell
        $token->{ref} = $self->decodeAreaRef2d($decoder);
      }
      elsif ($id == 0x1A)
      {
        # tRef3d - Reference to another cell on a different sheet
        $token->{ref} = $self->decodeCellRef3d($decoder);
      }
      else
      {
        ...;
      }
    }

    $formula->appendToken($token);
  }

  #print main::Dumper($tokens);

  return $formula;
}

1;
