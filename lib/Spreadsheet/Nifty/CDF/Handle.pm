#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::CDF::Handle;

use Fcntl qw();

use constant
{
  CACHE_SIZE_MAX => 4,
};

# === Class methods ===

sub new($$$)
{
  my $class = shift();
  my ($cdf, $chain, $size) = @_;

  my $self = {};
  $self->{open}      = 1;
  $self->{cdf}       = $cdf;
  $self->{chain}     = $chain;
  $self->{size}      = $size;
  $self->{position}  = 0;
  $self->{cache}     = {sectors => [], byIndex => {}};
  bless($self, $class);

  return $self;
}

# === Internal methods ===

sub _shrinkCache()
{
  my $self = shift();

  (scalar(@{$self->{cache}->{sectors}}) < CACHE_SIZE_MAX) && return;  # Already has room

  my $old = shift(@{$self->{cache}->{sectors}});
  delete($self->{cache}->{byIndex}->{$old->{index}});
  return;
}

# Caches a sector by this stream's 0-based chain index.
sub _cacheSector($)
{
  my $self = shift();
  my ($index) = @_;

  (defined($self->{cache}->{byIndex}->{$index})) && return;  # Already cached

  $self->_shrinkCache();

  my $data;
  if ($self->{chain}->{type} == Spreadsheet::Nifty::CDF::CHAIN_TYPE_MINI)
  {
    $data = $self->{cdf}->readMiniSector($self->{chain}->{sectors}->[$index]);
  }
  else
  {
    $data = $self->{cdf}->readSector($self->{chain}->{sectors}->[$index]);
  }

  my $new = {index => $index, data => $data};
  push(@{$self->{cache}->{sectors}}, $new);
  $self->{cache}->{byIndex}->{$index} = $new;
  return;
}

# === Instance methods ===

sub close()
{
  my $self = shift();

  $self->{open}     = 0;
  $self->{cdf}      = undef;
  $self->{chain}    = undef;
  $self->{size}     = undef;
  $self->{position} = undef;
  return;
}

sub opened()
{
  my $self = shift();

  return $self->{open};
}

sub eof()
{
  my $self = shift();

 return $self->{position} >= $self->{size};
}

sub tell()
{
  my $self = shift();

  return $self->{position};
}

sub seek($$)
{
  my $self = shift();
  my ($position, $whence) = @_;

  if ($whence == Fcntl::SEEK_CUR)
  {
    $position = $self->{position} + $position;
  }
  elsif ($whence == Fcntl::SEEK_END)
  {
    $position = $self->{size} + $position;
  }
  elsif ($whence != Fcntl::SEEK_SET)
  {
    die("seek(): Unknown whence value ${whence}");
  }

  $self->{position} = $position;
  return 1;
}

sub read($$)
{
  my $self = shift();
  my ($len) = $_[1];

  ($len <= 0) && return 0;  # Caller wants to read less than one byte

  # Don't try to read past the end of the stream
  if ($self->{position} + $len > $self->{size})
  {
    $len = $self->{size} - $self->{position};
    ($len <= 0) && return 0;  # No data left to read
  }

  my $data = '';
  my $j = int(($self->{position} + $len) / $self->{chain}->{sectorSize});
  while ($len > 0)
  {
    my $i = int($self->{position} / $self->{chain}->{sectorSize});
    $self->_cacheSector($i);

    my ($blkOffset, $blkLength, $blk);
    $blkOffset = $self->{position} % $self->{chain}->{sectorSize};

    if ($i < $j)
    {
      # Read the entire sector
      $blkLength = $self->{chain}->{sectorSize} - $blkOffset;
      $blk = substr($self->{cache}->{byIndex}->{$i}->{data}, $blkOffset);
    }
    else
    {
      # Read the amount needed from the final sector
      $blkLength = $len;
      $blk = substr($self->{cache}->{byIndex}->{$i}->{data}, $blkOffset, $blkLength);
    }

    $data .= $blk;
    $len -= $blkLength;
    $self->{position} += $blkLength;
  }
  
  $_[0] = $data;
  return length($data);
}

sub write()
{
  ...;
}

1;
