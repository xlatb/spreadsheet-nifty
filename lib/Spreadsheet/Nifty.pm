#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty;

use constant
{
  TYPE_NULL => 0x0,
  TYPE_NUM  => 0x1,
  TYPE_STR  => 0x2,
  TYPE_BOOL => 0x3,
  TYPE_ERR  => 0x4,
  TYPE_DATE => 0x5,
};

my $readers =
[
  {ext => qr#[.]xls$#i,     class => 'Spreadsheet::Nifty::XLS::FileReader'},
  {ext => qr#[.]xls[xm]$#i, class => 'Spreadsheet::Nifty::XLSX::FileReader'},
  {ext => qr#[.]xlsb$#i,    class => 'Spreadsheet::Nifty::XLSB::FileReader'},
  {ext => qr#[.]ods$#i,     class => 'Spreadsheet::Nifty::ODS::FileReader'},
];

# === Class methods ===

# Given a filename, returns a FileReader object or undef if the file is unsupported.
sub reader($)
{
  my $class = shift();
  my ($filename) = @_;

  for my $r (@{$readers})
  {
    ($filename !~ $r->{ext}) && next;

    my $readerPath = ($r->{class} =~ s#::#/#gr) . '.pm';
    require($readerPath);
    (!$r->{class}->isFileSupported($filename)) && next;

    my $reader = $r->{class}->new();
    return $reader;
  }

  return undef;
}

1;
