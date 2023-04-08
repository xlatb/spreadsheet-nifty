#!/usr/bin/perl -w
use warnings;
use strict;

# Read from an IO handle as a stream of BIFF records.
package Spreadsheet::Nifty::XLS::BIFFReader;

use Fcntl qw();
use Spreadsheet::Nifty::XLS::RecordTypes;

# === Class methods ===

sub new()
{
  my $class = shift();
  my ($io) = @_;

  my $self = {};
  $self->{io}     = $io;
  $self->{eof}    = 0;
  $self->{crypto} = undef;
  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub setDecryptor($)
{
  my $self = shift();
  my ($decryptor) = @_;

  $self->{crypto} = {offset => $self->{io}->tell(), decryptor => $decryptor};
  return;
}

# Copies the decryptor from another BIFFReader object.
sub copyDecryptor($)
{
  my $self = shift();
  my ($reader) = @_;

  if (!defined($reader->{crypto}))
  {
    $self->{crypto} = undef;
    return;
  }

  $self->{crypto} = {};
  $self->{crypto}->{offset} = $reader->{crypto}->{offset};
  $self->{crypto}->{decryptor} = $reader->{crypto}->{decryptor}->dup();
  return;
}

sub readRecordHeader()
{
  my $self = shift();

  my $header;
  my $count = $self->{io}->read($header, 4);
  if ($count == 0)
  {
    $self->{eof} = 1;
    return undef;
  }

  ($count < 0) && die("skipRecord(): Error reading record header: $!");
  ($count < 4) && die("readRecord(): Short read on record header");

  my ($type, $length) = unpack('vv', $header);

  return ($type, $length);
}

sub readRecordPayload($)
{
  my $self = shift();
  my ($type, $length) = @_;

  my $offset = $self->{io}->tell();

  my $payload;
  my $count = $self->{io}->read($payload, $length);
  ($count < 0) && die("readRecord(): Error reading record payload: $!");
  ($count < $length) && die(sprintf("readRecord(): Short read on record payload (expected %d bytes, got %d)", $length, $count));

  if (($count > 0) && defined($self->{crypto}) && ($offset >= $self->{crypto}->{offset}) && Spreadsheet::Nifty::XLS::Crypto->canRecordTypeBeEncrypted($type))
  {
    $payload = $self->{crypto}->{decryptor}->decryptBiffPayload($offset, $type, $count, $payload);
  }

  return $payload;
}

sub readRecord()
{
  my $self = shift();

  my $offset = $self->{io}->tell();

  my ($type, $length) = $self->readRecordHeader();
  (!defined($type)) && return undef;

  my $payload = $self->readRecordPayload($type, $length);

  my $name = Spreadsheet::Nifty::XLS::RecordTypes->name($type);  # TODO: Remove name?
  return {offset => $offset, type => $type, name => $name, length => $length, payload => $payload};
}

# If the next records has one of the given types, reads and returns it.
# Otherwise, returns undef and no record is consumed.
sub readRecordIfType($)
{
  my $self = shift();
  my ($expectedTypes) = @_;

  # Coerce scalar to array ref
  if (!ref($expectedTypes))
  {
    $expectedTypes = [ $expectedTypes ];
  }

  my $offset = $self->{io}->tell();

  my ($type, $length) = $self->readRecordHeader();
  (!defined($type)) && return undef;

  if (!grep({ $type == $_} @{$expectedTypes}))
  {
    # Not any of the expected types, jump back to start of record
    $self->seek($offset);
    return undef;
  }

  my $payload = $self->readRecordPayload($type, $length);

  my $name = Spreadsheet::Nifty::XLS::RecordTypes->name($type);  # TODO: Remove name?
  return {offset => $offset, type => $type, name => $name, length => $length, payload => $payload};  
}

# If the next record has one of the given types, skips over it and returns true.
# Otherwise, returns false and no record is consumed.
sub skipRecordIfType($)
{
  my $self = shift();
  my ($expectedTypes) = @_;

  # Coerce scalar to array ref
  if (!ref($expectedTypes))
  {
    $expectedTypes = [ $expectedTypes ];
  }

  my $offset = $self->{io}->tell();

  my ($type, $length) = $self->readRecordHeader();
  (!defined($type)) && return 0;

  #if ($type != $expectedType)
  if (!grep({ $type == $_} @{$expectedTypes}))
  {
    # Not any of the expected types, jump back to start of record
    $self->seek($offset);
    return 0;
  }

  # Jump over the payload
  $self->seek($offset + 4 + $length);
  return 1;
}

sub tell()
{
  my $self = shift();

  return $self->{io}->tell();
}

sub seek($)
{
  my $self = shift();
  my ($pos) = @_;

  (!defined($pos)) && die("Seek to undefined position");

  return $self->{io}->seek($pos, Fcntl::SEEK_SET);
}

1;
