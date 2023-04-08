#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLS::Crypto;

sub rc4(@)
{
  my $class = shift();

  require Spreadsheet::Nifty::XLS::Crypto::RC4;
  return Spreadsheet::Nifty::XLS::Crypto::RC4->new(@_);
}

sub xor(@)
{
  my $class = shift();

  require Spreadsheet::Nifty::XLS::Crypto::Xor;
  return Spreadsheet::Nifty::XLS::Crypto::Xor->new(@_);
}

sub cryptoApiRC4(@)
{
  my $class = shift();

  require Spreadsheet::Nifty::XLS::Crypto::CryptoApiRC4;
  return Spreadsheet::Nifty::XLS::Crypto::CryptoApiRC4->new(@_);
}

# Returns true iff BIFF records of the given type can be encrypted.
sub canRecordTypeBeEncrypted($)
{
  my $class = shift();
  my ($type) = @_;

  ($type == 0x0809) && return 0;  # BOF
  ($type == 0x002F) && return 0;  # FilePass
  ($type == 0x0194) && return 0;  # UsrExcl
  ($type == 0x0195) && return 0;  # FileLock
  ($type == 0x00E1) && return 0;  # InterfaceHdr
  ($type == 0x0196) && return 0;  # RRDInfo
  ($type == 0x0138) && return 0;  # RRDHead

  return 1;
}

1;
