#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::CDF::DirectoryEntry;

# === Class methods ===

sub new($$)
{
  my $class = shift();
  my ($cdf, $entry) = @_;

  my $self = {};
  $self->{cdf}   = $cdf;
  $self->{entry} = $entry;
  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub name()
{
  my $self = shift();

  return $self->{entry}->{name};
}

sub type()
{
  my $self = shift();

  return $self->{entry}->{type};
}

sub isStream()
{
  my $self = shift();

  return ($self->{entry}->{type} == Spreadsheet::Nifty::CDF::OBJ_TYPE_STREAM);
}

sub isDirectory()
{
  my $self = shift();

  return ($self->{entry}->{type} == Spreadsheet::Nifty::CDF::OBJ_TYPE_DIR) || ($self->{entry}->{type} == Spreadsheet::Nifty::CDF::OBJ_TYPE_ROOTDIR);
}

sub rawClassId()
{
  my $self = shift();

  return $self->{entry}->{classId};
}

sub classId()
{
  my $self = shift();

  return Spreadsheet::Nifty::CDF->formatUUID($self->{entry}->{classId});
}

sub startSector()
{
  my $self = shift();

  ($self->{entry}->{type} != Spreadsheet::Nifty::CDF::OBJ_TYPE_STREAM) && return undef;
  return $self->{entry}->{startSector};
}

sub size()
{
  my $self = shift();

  ($self->{entry}->{type} != Spreadsheet::Nifty::CDF::OBJ_TYPE_STREAM) && return undef;
  return $self->{entry}->{size};
}

sub createTime()
{
  my $self = shift();

  return Spreadsheet::Nifty::CDF::winTimeToPosixTime($self->{entry}->{createTime});
}

sub modifyTime()
{
  my $self = shift();

  return Spreadsheet::Nifty::CDF::winTimeToPosixTime($self->{entry}->{modifyTime});
}

sub firstChild()
{
  my $self = shift();

  ($self->{entry}->{childId} == Spreadsheet::Nifty::CDF::DIR_INDEX_NONE) && return undef;

  return $self->{cdf}->getDirectoryEntry($self->{entry}->{childId});
}

# Returns the set of sibling directory entries that contains this directory entry.
sub siblings()
{
  my $self = shift();

  my $siblings = [];

  if ($self->{entry}->{leftId} != Spreadsheet::Nifty::CDF::DIR_INDEX_NONE)
  {
    my $left = $self->{cdf}->getDirectoryEntry($self->{entry}->{leftId});
    push(@{$siblings}, @{$left->siblings()});
  }

  push(@{$siblings}, $self);

  if ($self->{entry}->{rightId} != Spreadsheet::Nifty::CDF::DIR_INDEX_NONE)
  {
    my $right = $self->{cdf}->getDirectoryEntry($self->{entry}->{rightId});
    push(@{$siblings}, @{$right->siblings()});
  } 

  return $siblings;
}

# Returns the child directory entries of this directory entry.
sub children()
{
  my $self = shift();

  ($self->{entry}->{childId} == Spreadsheet::Nifty::CDF::DIR_INDEX_NONE) && return [];

  my $child = $self->{cdf}->getDirectoryEntry($self->{entry}->{childId});
  return $child->siblings();
}

sub getChildNamed($)
{
  my $self = shift();
  my ($name) = @_;

  ($self->{entry}->{childId} == Spreadsheet::Nifty::CDF::DIR_INDEX_NONE) && return undef;

  my $children = $self->children();
  for my $c (@{$children})
  {
    ($c->name() eq $name) && return $c;
  }

  return undef;
}

sub getChain()
{
  my $self = shift();

  ($self->{entry}->{type} != Spreadsheet::Nifty::CDF::OBJ_TYPE_STREAM) && return undef;

  my $chain = {};
  if ($self->{entry}->{size} < $self->{cdf}->{header}->{miniStreamMaxObjectSize})
  {
    $chain->{type} = Spreadsheet::Nifty::CDF::CHAIN_TYPE_MINI;
    $chain->{sectors} = $self->{cdf}->getMiniFatChain($self->{entry}->{startSector});
    $chain->{sectorSize} = $self->{cdf}->{miniSectorSize};
  }   
  else
  {
    $chain->{type} = Spreadsheet::Nifty::CDF::CHAIN_TYPE_NORMAL;
    $chain->{sectors} = $self->{cdf}->getFatChain($self->{entry}->{startSector});
    $chain->{sectorSize} = $self->{cdf}->{sectorSize};
  }

  return $chain;
}

sub getBytes()
{
  my $self = shift();

  ($self->{entry}->{type} != Spreadsheet::Nifty::CDF::OBJ_TYPE_STREAM) && return undef;

  my $bytes;
  if ($self->{entry}->{size} < $self->{cdf}->{header}->{miniStreamMaxObjectSize})
  {
    $bytes = $self->{cdf}->readMiniSectorChain($self->{cdf}->getMiniFatChain($self->{entry}->{startSector}));
  }
  else
  {
    $bytes = $self->{cdf}->readSectorChain($self->{cdf}->getFatChain($self->{entry}->{startSector}));
  }

  return $bytes;
}

sub open($)
{
  my $self = shift();
  my ($mode) = @_;

  ($self->{entry}->{type} != Spreadsheet::Nifty::CDF::OBJ_TYPE_STREAM) && return undef;

  $mode //= 'r';
  ($mode ne 'r') && do { ...; };

  my $chain = $self->getChain();
  my $handle = Spreadsheet::Nifty::CDF::Handle->new($self->{cdf}, $chain, $self->{entry}->{size});
  return $handle;
}

1;
