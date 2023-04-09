#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty;

my $readers =
[
  {ext => qr#[.]xls$#i,  class => 'Spreadsheet::Nifty::XLS::FileReader'},
  {ext => qr#[.]xlsb$#i, class => 'Spreadsheet::Nifty::XLSB::FileReader'},
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
