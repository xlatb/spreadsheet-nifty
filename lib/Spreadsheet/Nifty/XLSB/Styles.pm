#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSB::Styles;

use Spreadsheet::Nifty::XLSB::RecordReader;
use Spreadsheet::Nifty::XLSB::RecordTypes;
use Spreadsheet::Nifty::XLSB::Decode;
use Spreadsheet::Nifty::StructDecoder;
use Spreadsheet::Nifty::IndexedColors;

# === Class methods ===

sub new()
{
  my $class = shift();

  my $self = {};
  $self->{debug}         = 0;
  $self->{numberFormats} = {};
  $self->{fonts}         = [];
  $self->{fills}         = [];
  $self->{borders}       = [];
  $self->{cellStyles}    = [];
  $self->{styles}        = [];
  $self->{xfs}           = [];
  $self->{palette}       = {indexed => Spreadsheet::Nifty::IndexedColors->new(), mru => []};
  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub read($)
{
  my $self = shift();
  my ($reader) = @_;

  my $recs = Spreadsheet::Nifty::XLSB::RecordReader->new($reader);
  while (my $rec = $recs->read())
  {
    #printf("REC type %d size %d name %s\n", $rec->{type}, $rec->{size}, $rec->{name} // '?');
    #($rec->{size}) && printf("  data: %s\n", unpack('H*', $rec->{data}));

    if ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::BEGIN_FMTS)
    {
      $self->readNumberFormats($recs) || return !!0;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::BEGIN_FONTS)
    {
      $self->readFonts($recs) || return !!0;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::BEGIN_FILLS)
    {
      $self->readFills($recs) || return !!0;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::BEGIN_BORDERS)
    {
      $self->readBorders($recs) || return !!0;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::BEGIN_CELL_STYLE_XFS)
    {
      $self->readCellStyles($recs) || return !!0;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::BEGIN_STYLES)
    {
      $self->readStyles($recs) || return !!0;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::BEGIN_CELL_XFS)
    {
      $self->readCellXfs($recs) || return !!0;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::BEGIN_COLOR_PALETTE)
    {
      $self->readColorPalette($recs) || return !!0;
    }
  }

  return !!1;
}

sub readNumberFormats($)
{
  my $self = shift();
  my ($recs) = @_;

  while (my $rec = $recs->read())
  {
    #printf("REC type %d size %d name %s\n", $rec->{type}, $rec->{size}, $rec->{name} // '?');
    #($rec->{size}) && printf("  data: %s\n", unpack('H*', $rec->{data}));

    if ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::END_FMTS)
    {
      return !!1;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::FMT)
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $fmt = $decoder->decodeHash(['id:u16', 'string:XLWideString']);
      $self->{numberFormats}->{$fmt->{id}} = $fmt->{string};
      #use Data::Dumper; print Dumper($fmt);
    }
  }

  return !!0;  # Hit end of records before seeing the end-of-number-formats record
}

sub readFonts($)
{
  my $self = shift();
  my ($recs) = @_;

  my $fonts = [];

  while (my $rec = $recs->read())
  {
    if ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::END_FONTS)
    {
      $self->{fonts} = $fonts;
      return !!1;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::FONT)
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $font = $decoder->decodeHash(['height:u16', 'flags:u16', 'weight:u16', 'supersub:u16', 'underline:u8', 'family:u8', 'charset:u8', ':u8', 'color:BrtColor', 'scheme:u8', 'name:XLWideString']);
      push(@{$fonts}, $font);
    }
  }

  return !!0;  # Hit end of records before seeing the end-of-fonts record
}

sub readFills($)
{
  my $self = shift();
  my ($recs) = @_;

  my $fills = [];

  while (my $rec = $recs->read())
  {
    if ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::END_FILLS)
    {
      $self->{fills} = $fills;
      return !!1;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::FILL)
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $fill = $decoder->decodeHash(['pattern:u32', 'fgColor:BrtColor', 'bgColor:BrtColor']);
      $fill->{gradient} = $decoder->decodeHash(['type:u32', 'angle:f64', 'fillLeft:f64', 'fillRight:f64', 'fillTop:f64', 'fillBottom:f64', 'stopCount:u32']);
      $fill->{gradient}->{stops} = $decoder->decodeArray(['color:BrtColor', 'position:f64'], $fill->{gradient}->{stopCount});
      push(@{$fills}, $fill);
    }
  }

  return !!0;  # Hit end of records before seeing the end-of-fills record
}

sub readBorders($)
{
  my $self = shift();
  my ($recs) = @_;

  my $borders = [];

  while (my $rec = $recs->read())
  {
    if ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::END_BORDERS)
    {
      $self->{borders} = $borders;
      return !!1;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::BORDER)
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $border = $decoder->decodeHash(['flags:u8', 'top:Blxf', 'bottom:Blxf', 'left:Blxf', 'right:Blxf', 'diagonal:Blxf']);
      push(@{$borders}, $border);
    }
  }

  return !!0;  # Hit end of records before seeing the end-of-borders record
}

sub readCellStyles($)
{
  my $self = shift();
  my ($recs) = @_;

  my $cellStyles = [];

  while (my $rec = $recs->read())
  {
    if ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::END_CELL_STYLE_XFS)
    {
      $self->{cellStyles} = $cellStyles;
      return !!1;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::XF)
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $xf = $decoder->decodeField('BrtXF');
      push(@{$cellStyles}, $xf);
    }
  }

  return !!0;  # Hit end of records before seeing the end-of-cell-styles record
}

sub readStyles($)
{
  my $self = shift();
  my ($recs) = @_;

  my $styles = [];

  while (my $rec = $recs->read())
  {
    if ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::END_STYLES)
    {
      $self->{styles} = $styles;
      return !!1;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::STYLE)
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $style = $decoder->decodeHash(['xfId:u32', 'flags:u16', 'builtin:u16', 'name:XLNullableWideString']);
      push(@{$styles}, $style);
    }
  }

  return !!0;  # Hit end of records before seeing the end-of-styles record
}

sub readCellXfs($)
{
  my $self = shift();
  my ($recs) = @_;

  my $xfs = [];

  while (my $rec = $recs->read())
  {
    if ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::END_CELL_XFS)
    {
      $self->{xfs} = $xfs;
      return !!1;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::XF)
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $xf = $decoder->decodeField('BrtXF');
      push(@{$xfs}, $xf);
    }
  }

  return !!0;  # Hit end of records before seeing the end-of-cell-xfs record
}

sub readColorPalette($)
{
  my $self = shift();
  my ($recs) = @_;

  my $palette = {};
  $palette->{indexed} = Spreadsheet::Nifty::IndexedColors->new();
  $palette->{mru}     = [];

  while (my $rec = $recs->read())
  {
    if ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::END_COLOR_PALETTE)
    {
      $self->{palette} = $palette;
      return !!1;
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::BEGIN_INDEXED_COLORS)
    {
      while (my $rec = $recs->read())
      {
        ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::END_INDEXED_COLORS) && last;

        if ($rec->{type} == $rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::INDEXED_COLOR)
        {
          my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
          my $c = $decoder->decodeHash(['r:u8', 'g:u8', 'b:u8']);
          $palette->{indexed}->addColorRGB($c);
        }
      }
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::BEGIN_MRU_COLORS)
    {
      while (my $rec = $recs->read())
      {
        ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::END_MRU_COLORS) && last;

        my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
        my $color = $decoder->decodeField('BrtColor');
        push(@{$palette->{mru}}, $color);
      }
    }
  }

  return !!0;  # Hit end of records before seeing the end-of-color-palette record
}

sub getXf($)
{
  my $self = shift();
  my ($i) = @_;

  (($i < 0) || ($i >= scalar(@{$self->{xfs}}))) && return undef;  # Out of range

  my $xf = $self->{xfs}->[$i];
  return $xf;
}

sub getFill($)
{
  my $self = shift();
  my ($i) = @_;

  (($i < 0) || ($i >= scalar(@{$self->{fills}}))) && return undef;  # Out of range

  my $fill = $self->{fills}->[$i];
  return $fill;
}

sub getNumberFormat($)
{
  my $self = shift();
  my ($i) = @_;

  my $fmt = $self->{numberFormats}->{$i};
  return $fmt;
}

1;
