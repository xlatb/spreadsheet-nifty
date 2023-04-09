#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSB::RecordReader;

# === Class methods ===

sub new()
{
  my $class = shift();
  my ($reader) = @_;

  my $self = {};
  $self->{reader} = $reader;

  bless($self, $class);

  return $self;
}

# === Instance methods ===

# Bytes are stored low to high, with the high bit telling us whether the value continues.
sub readVariableInt($)
{
  my $self = shift();
  my ($maxsize) = @_;

  my $size = 0;
  my $v = 0;

  my $buf;

  while (1)
  {
    my $count = $self->{reader}->read($buf, 1);
    ($count == 0) && return undef;

    my $b = ord($buf);
    #printf("B: %02X v: %02X\n", $b, $v);
    $v |= (($b & 0x7F) << (7 * $size));
    ($b & 0x80) || return $v;

    $size++;
    (defined($maxsize)) && ($size == $maxsize) && return $v;
  }
}

sub readBytes($)
{
  my $self = shift();
  my ($size) = @_;

  my $buf;
  my $count = $self->{reader}->read($buf, $size);
  ($count < $size) && die("Short read");

  return $buf;
}

sub read($)
{
  my $self = shift();

  my $type = $self->readVariableInt(2);
  (!defined($type)) && return undef;

  my $size = $self->readVariableInt(4);
  (!defined($size)) && return undef;

  my $data = $self->readBytes($size);
  (!defined($data)) && return undef;

  my $name = Spreadsheet::Nifty::XLSB::RecordTypes->name($type);

  return {type => $type, size => $size, data => $data, name => $name};
}

1;
