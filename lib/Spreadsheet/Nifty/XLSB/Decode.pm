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

1;
