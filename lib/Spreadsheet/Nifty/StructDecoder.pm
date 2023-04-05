#!/usr/bin/perl -w
use warnings;
use strict;

package StructDecoder;

use Encode qw();

# === Class methods ===

sub new()
{
  my $class = shift();
  my ($bytes) = @_;

  (ref($bytes)) && die("Expected byte string");

  my $self = {};
  $self->{bytes}    = $bytes;
  $self->{position} = 0;
  $self->{length}   = length($bytes);
  $self->{types}    = {};

  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub registerType($$)
{
  my $self = shift();
  my ($name, $decoder) = @_;

  $self->{types}->{$name} = {decoder => $decoder};
}

sub assertLength($)
{
  my $self = shift();
  my ($len) = @_;

  my $avail = $self->{length} - $self->{position};
  ($avail < $len) && die("Wanted $len bytes but there are only $avail available");
}

sub bytesLeft()
{
  my $self = shift();

  return $self->{length} - $self->{position};
}

sub skipBytes($)
{
  my $self = shift();
  my ($len) = @_;

  $self->{position} += $len;
  if ($self->{position} > $self->{length})
  {
    $self->{position} = $self->{length};
  }

  return;
}

sub getBytes($)
{
  my $self = shift();
  my ($len) = @_;

  my $avail = $self->{length} - $self->{position};
  ($avail < $len) && die("Wanted $len bytes but there are only $avail available");

  my $bytes = substr($self->{bytes}, $self->{position}, $len);
  $self->{position} += $len;
  return $bytes;
}

sub setBytes($)
{
  my $self = shift();
  my ($bytes) = @_;

  $self->{bytes}    = $bytes;
  $self->{position} = 0;
  $self->{length}   = length($bytes);

  return;
}

sub appendBytes($)
{
  my $self = shift();
  my ($bytes) = @_;

  $self->{bytes} .= $bytes;
  $self->{length} = length($self->{bytes});

  return;
}

sub decodeFieldInternal($;$)
{
  my $self = shift();
  my ($type, $size) = @_;

  if ($type eq 'u8')
  {
    return ord($self->getBytes(1));
  }
  elsif ($type eq 'u16')
  {
    my $v = unpack('v', $self->getBytes(2));
    return $v;
  }
  elsif ($type eq 'u24')
  {
    my $bytes = $self->getBytes(3);
    return (ord(substr($bytes, 2, 1)) << 16) | (ord(substr($bytes, 1, 1)) << 8) | ord(substr($bytes, 0, 1));
  }
  elsif ($type eq 'u32')
  {
    my $v = unpack('V', $self->getBytes(4));
    return $v;
  }
  elsif ($type eq 'u64')
  {
    my $v = unpack('Q<', $self->getBytes(8));
    return $v;
  }
  elsif ($type eq 'f32')
  {
    my $v = unpack('f<', $self->getBytes(4));
    return $v;
  }
  elsif ($type eq 'f64')
  {
    my $v = unpack('d<', $self->getBytes(8));
    return $v;
  }
  elsif ($type eq 'bytes')
  {
    my $v = $self->getBytes($size);
    return $v;
  }
  elsif (defined(my $t = $self->{types}->{$type}))
  {
    return $t->{decoder}->($self, $size);
  }
  else
  {
    die("decodeField(): Unknown type '$type'")
  }
}

sub decodeField($)
{
  my $self = shift();
  my ($type) = @_;

  my $size;
  if ($type =~ m#^([a-zA-Z]\w+)\[([^]]+)\]$#)
  {
    $type = $1;
    $size = $2;
  }

  return $self->decodeFieldInternal($type, $size);
}

sub decodeHash($)
{
  my $self = shift();
  my ($fields) = @_;

  my $hash = {};

  for my $field (@{$fields})
  {
    ($field !~ m#^(\w*)(?:\[(\d+)\])?:(\w+)(?:\[(\w+)\])?$#) && die("Bad field spec: '$field'");
    my ($name, $count, $type, $size) = ($1, $2, $3, $4);
    #STDERR->printf("SPEC: '%s' name '%s' count '%s', type '%s', size '%s'\n", $field, $name // 'undef', $count // 'undef', $type // 'undef', $size // 'undef');

    # If the type had a size, interpret it
    if (defined($size))
    {
      if ($size =~ m#^\d+$#)
      {
        $size = int($size);
      }
      elsif (defined($hash->{$size}))
      {
        $size = $hash->{$size};
      }
      else
      {
        die("Unknown size spec: '$size'");
      }
      #STDERR->printf("SIZE: '%s'\n", $size);
    }

    (defined($count)) && (defined($size)) && die("Array with count of type with size not supported");  # TODO
    my $val = defined($count) ? $self->decodeArray($type, int($count)) : $self->decodeFieldInternal($type, $size);

    (length($name)) && do { $hash->{$name} = $val; };
  }

  return $hash;
}

sub decodeArray($$)
{
  my $self = shift();
  my ($fields, $count) = @_;

  my $array = [];

  if (!ref($fields))
  {
    for (my $i = 0; $i < $count; $i++)
    {
      push(@{$array}, $self->decodeField($fields));
    }
  }
  else
  {
    for (my $i = 0; $i < $count; $i++)
    {
      my $struct = $self->decodeHash($fields);
      push(@{$array}, $struct);
    }
  }

  return $array;
}

1;
