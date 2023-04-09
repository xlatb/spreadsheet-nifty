#!/usr/bin/perl -w
use warnings;
use strict;

# A wrapper for Archive::Zip::Member to provide an IO::Handle-like interface.

package Spreadsheet::Nifty::ZIPReader;

use Archive::Zip qw(:ERROR_CODES :CONSTANTS);
use Fcntl qw();

# === Class methods ===

sub new()
{
  my $class = shift();
  my ($member) = @_;
  
  my $self = {};
  $self->{open}          = 0;
  $self->{zipStreamDone} = 1;
  $self->{member}        = $member;
  $self->{buffer}        = '';
  
  bless($self, $class);
  
  $self->_open();
  
  return $self;
}

sub new_from_fd()
{
  my $class = shift();
  
  die($class . ": Unsupported: new_from_fd()");
}

# === Instance methods ===

sub _open()
{
  my $self = shift();

  $self->{member}->desiredCompressionMethod(COMPRESSION_STORED);
  my $status = $self->{member}->rewindData();
  
  if ($status == AZ_OK)
  {
    $self->{open} = 1;
    $self->{zipStreamDone} = 0;
  }
  
  return;
}

sub _close()
{
  my $self = shift();

  $self->{member}->endRead();
  
  $self->{open} = 0;
  $self->{zipStreamDone} = 1;
  
  return;
}

sub close()
{
  my $self = shift();
  
  $self->_close();
  
  return;
}

sub read($$$)
{
  my $self = shift();
  my ($buf, $len, $offset) = @_;
  
  my $bufref = \$_[0];
  
  (defined($offset)) && die(ref($self) . ": Unsupported: read() with offset");

  while (1)
  {
    # If we can satisfy the read entirely from the buffer, do it
    if ($len <= length($self->{buffer}))
    {
      $$bufref = substr($self->{buffer}, 0, $len);
      $self->{buffer} = substr($self->{buffer}, $len);
      return $len;
    }
    
    if (!$self->{zipStreamDone})
    {
      # Read from the ZIP member stream
      my ($zipDataRef, $status) = $self->{member}->readChunk();
      if ($status == AZ_STREAM_END)
      {
        $self->{zipStreamDone} = 1;
      }
      elsif ($status != AZ_OK)
      {
        return undef;  # Signal error
      }
      
      $self->{buffer} .= $$zipDataRef;
    }
    else
    {
      # Not enough data buffered, and ZIP stream is done, give back as much data as possible
      $$bufref = $self->{buffer};
      $len = length($self->{buffer});
      $self->{buffer} = '';
      return $len;
    }
  }
}

sub seek($;$)
{
  my $self = shift();
  my ($position, $whence) = @_;

  $whence //= Fcntl::SEEK_SET;

  ($whence != Fcntl::SEEK_SET) && die(ref($self) . ": Unsupported: seek() must use SEEK_SET");

  my $status = $self->{member}->rewindData();
  ($status != AZ_OK) && return 0;

  $self->{zipStreamDone} = 0;

  ($position == 0) && return 1;  # All done

  die(ref($self) . ": Unsupported: seek() can only return to position 0");
}

1;
