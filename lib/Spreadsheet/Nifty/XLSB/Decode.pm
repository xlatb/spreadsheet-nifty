#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSB::Decode;

use Spreadsheet::Nifty::Utils;
use Spreadsheet::Nifty::StructDecoder;

use Encode qw();

sub decoder()
{
  my $class = shift();

  my $decoder = StructDecoder->new(@_);
  $decoder->registerType('XLWideString', \&decodeXLWideString);
  $decoder->registerType('XLNullableWideString', \&decodeXLNullableWideString);
  $decoder->registerType('BrtColor', \&decodeBrtColor);
  $decoder->registerType('BrtXF', \&decodeBrtXF);
  $decoder->registerType('Blxf', \&decodeBlxf);
  return $decoder;
}

# === Single-field decoders ===

sub decodeXLWideString()
{
  my $decoder = shift();

  my $size = unpack('V', $decoder->getBytes(4));
  my $payload = $decoder->getBytes($size * 2);

  my $str = Encode::decode('UTF-16LE', $payload);
  return $str;
}

sub decodeXLNullableWideString()
{
  my $decoder = shift();

  my $size = unpack('V', $decoder->getBytes(4));
  ($size == 0xFFFFFFFF) && return undef;  # Null string

  my $payload = $decoder->getBytes($size * 2);

  my $str = Encode::decode('UTF-16LE', $payload);
  return $str;
}

sub decodeBrtColor()
{
  my $decoder = shift();

  return $decoder->decodeHash(['type:u8', 'index:u8', 'tint:i16', 'red:u8', 'green:u8', 'blue:u8', 'alpha:u8']);
}

sub decodeBrtXF()
{
  my $decoder = shift();

  return $decoder->decodeHash(['parentStyle:u16', 'numberFormatId:u16', 'fontId:u16', 'fillId:u16', 'borderId:u16', 'textRotation:u8', 'indent:u8', 'flags:u32']);
}

sub decodeBlxf()
{
  my $decoder = shift();

  return $decoder->decodeHash(['type:u8', ':u8', 'color:BrtColor']);
}

1;
