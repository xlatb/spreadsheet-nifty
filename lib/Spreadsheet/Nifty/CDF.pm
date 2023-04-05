#!/usr/bin/perl -w
use warnings;
use strict;

# https://en.wikipedia.org/wiki/Compound_File_Binary_Format

package Spreadsheet::Nifty::CDF;

use constant
{
  CHAIN_TYPE_MINI   => 0,
  CHAIN_TYPE_NORMAL => 1,
};

use constant
{
  SECTOR_FREE         => 0xFFFFFFFF,
  SECTOR_END_OF_CHAIN => 0xFFFFFFFE,
  SECTOR_FAT          => 0xFFFFFFFD,
  SECTOR_DIFAT        => 0xFFFFFFFC,
  SECTOR_MAX          => 0xFFFFFFF9,
};

use constant
{
  DIR_INDEX_NONE => 0xFFFFFFFF,
  DIR_INDEX_MAX  => 0xFFFFFFFA,
};

use constant
{
  DIFAT_INLINE_SIZE => 109,  # Number of DiFat entries included within the header

  DIRECTORY_ENTRY_SIZE => 128,

  MINIFAT_ENTRY_SIZE => 4,
};

use constant
{
  OBJ_TYPE_NONE    => 0,
  OBJ_TYPE_DIR     => 1,
  OBJ_TYPE_STREAM  => 2,
  OBJ_TYPE_ROOTDIR => 5,
};

use Spreadsheet::Nifty::CDF::DirectoryEntry;
use Spreadsheet::Nifty::CDF::Handle;
use Spreadsheet::Nifty::StructDecoder;

use Encode;
use Fcntl qw();

# === Class methods ===

sub new($)
{
  my $class = shift();
  my ($io) = @_;

  my $self = {};
  $self->{io} = $io;
  $self->{sectorSize} = undef;
  $self->{miniSectorSize} = undef;
  $self->{header} = undef;
  $self->{diFatMap} = undef;
  $self->{objectCount} = undef;
  bless($self, $class);

  return $self;
}

# Converts a 64-bit Windows FILETIME to a POSIX time value.
sub winTimeToPosixTime($)
{
  my $class = shift();
  my ($winTime) = @_;

  my $time = $winTime - 134774 * 24 * 3600;  # 134774 days from 1601-01-01 to 1970-01-01
  $time /= 10000;  # 10000 ticks per second to 1 tick per second

  return $time;
}

# Formats a 16-byte UUID value as a string.
sub formatUUID($)
{
  my $class = shift();
  my ($uuid) = @_;

  (length($uuid) != 16) && die("Bad UUID length");

  my $variant = ord(substr($uuid, 8, 1)) >> 4;
  if (($variant == 0xC) || ($variant == 0xD))
  {
    # Microsoft byte order
    my @parts = unpack('Vvv', $uuid);
    return sprintf("%08X-%04X-%04X-%s-%s", $parts[0], $parts[1], $parts[2], lc(unpack('H*', substr($uuid, 8, 2))), lc(unpack('H*', substr($uuid, 10, 6))));
  }
  else
  {
    # Standard byte order
    my @parts = (substr($uuid, 0, 4), substr($uuid, 4, 2), substr($uuid, 6, 2), substr($uuid, 8, 2), substr($uuid, 10, 6));
    return sprintf("%s-%s-%s-%s-%s", map({ lc(unpack('H*', $_)) } @parts));
  }
}

# === Instance methods ===

sub open()
{
  my $self = shift();

  ($self->readHeader()) || return 0;

  ($self->readDiFatMap()) || return 0;

  ($self->readDirectory()) || return 0;

  return 1;
}

sub readHeader()
{
  my $self = shift();

  (!$self->{io}->sysseek(0, Fcntl::SEEK_SET)) && return undef;  # Can't seek

  my $block;
  my $count = $self->{io}->sysread($block, 512);
  ($count < 512) && return undef;  # Short read

  my $decoder = StructDecoder->new($block);

  my $signature = $decoder->decodeField('bytes[8]');
  ($signature ne "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1") && return undef;  # Signature mismatch

  my $header = $decoder->decodeHash(
  [
    'classId:bytes[16]',
    'versionMinor:u16',
    'versionMajor:u16',
    'byteOrder:u16',
    'sectorShift:u16',
    'miniSectorShift:u16',
    ':bytes[6]',
    'directorySectorCount:u32',
    'fatSectorCount:u32',
    'directoryStartSector:u32',
    'transactionNumber:u32',
    'miniStreamMaxObjectSize:u32',
    'miniFatStartSector:u32',
    'miniFatSectorCount:u32',
    'diFatStartSector:u32',
    'diFatSectorCount:u32',
    'diFatInline[' . DIFAT_INLINE_SIZE . ']:u32',
  ]);

  $self->{header} = $header;

  $self->{sectorSize}     = 1 << $self->{header}->{sectorShift};
  $self->{miniSectorSize} = 1 << $self->{header}->{miniSectorShift};

  return 1;
}

sub readDiFatMap()
{
  my $self = shift();

  my $map = [];

  my $nextSector = $self->{header}->{diFatStartSector};
  my $seenSector = {};
  while ($nextSector != SECTOR_END_OF_CHAIN)
  {
    #printf("readDiFatMap(): nextSector %d (0x%08X)\n", $nextSector, $nextSector);
    # Chain should end with SECTOR_END_OF_CHAIN (0xFFFFFFFE), but some test
    #  files have used SECTOR_FREE (0xFFFFFFFF) instead. We allow the latter
    #  value here when the number of sectors in the chain is consistent with
    #  what we expect from the header.
    ($nextSector == SECTOR_FREE) && (scalar(@{$map}) == $self->{header}->{diFatSectorCount}) && last;

    defined($seenSector->{$nextSector}) && die("readDiFatMap(): Loop detected in DiFat chain");
    ($nextSector > SECTOR_MAX) && die("readDiFatMap(): Unexpected next sector number ${nextSector}");

    push(@{$map}, $nextSector);
    $seenSector->{$nextSector} = 1;

    my $sector = $self->readSector($nextSector);
    (!defined($sector)) && die("readDiFatMap(): Failed to read sector $nextSector");

    $nextSector = unpack('V', substr($sector, -4, 4));
  }

  $self->{diFatMap} = $map;
  return 1;
}

sub readDirectory()
{
  my $self = shift();

  my $chain = $self->getFatChain($self->{header}->{directoryStartSector});
  my $bytes = $self->readSectorChain($chain);
  (!defined($bytes)) && die("readDirectory(): Could not read directory sectors");

  my $directory = [];
  my $decoder = StructDecoder->new($bytes);
  my $count = length($bytes) / DIRECTORY_ENTRY_SIZE;
  for (my $i = 0; $i < $count; $i++)
  {
    my $entry = $decoder->decodeHash(
    [
      'name:bytes[64]',
      'nameLength:u16',  # Bytes, not characters, including trailing UTF-16 NUL
      'type:u8',
      'color:u8',
      'leftId:u32',
      'rightId:u32',
      'childId:u32',
      'classId:bytes[16]',
      'flags:u32',
      'createTime:bytes[8]',
      'modifyTime:bytes[8]',
      'startSector:u32',
      'size:u64',
    ]);

    $entry->{name} = Encode::decode('UTF-16LE', substr($entry->{name}, 0, $entry->{nameLength} - 2));

    ($self->{header}->{versionMajor} == 3) && do { $entry->{size} &= 0x7FFFFFFF; };

    push(@{$directory}, $entry);
  }

  $self->{directory} = $directory;
  return 1;
}

sub readSector($)
{
  my $self = shift();
  my ($s) = @_;

  ($s < 0) && die("readSector(): Negative sector number");
  my $offset = ($s + 1) * $self->{sectorSize};
  (!$self->{io}->sysseek($offset, Fcntl::SEEK_SET)) && return undef;  # Can't seek

  my $block;
  my $count = $self->{io}->sysread($block, $self->{sectorSize});
  ($count < $self->{sectorSize}) && return undef;  # Short read

  return $block;
}

sub readSectorChain($)
{
  my $self = shift();
  my ($chain) = @_;

  my $bytes = '';

  for (my $i = 0; $i < scalar(@{$chain}); $i++)
  {
    # TODO: Combine reads?
    my $sector = $self->readSector($chain->[$i]);
    (!defined($sector)) && return undef;
    $bytes .= $sector;
  }

  return $bytes;
}

# Read a mini sector from the mini stream.
sub readMiniSector($)
{
  my $self = shift();
  my ($ms) = @_;

  ($ms < 0) && die("readMiniSector(): Negative mini sector number");
  my $miniStreamOffset = $ms * $self->{miniSectorSize};

  my $fatSectorIndex = int($miniStreamOffset / $self->{sectorSize});
  my $fatSectorOffset = $miniStreamOffset % $self->{sectorSize};

  my $miniStreamFatChain = $self->getFatChain($self->{directory}->[0]->{startSector});
  my $sector = $self->readSector($miniStreamFatChain->[$fatSectorIndex]);
  (!defined($sector)) && die("readMiniSector(): Could not read mini sector index ${ms} sector index ${fatSectorIndex}");

  my $bytes = substr($sector, $fatSectorOffset, $self->{miniSectorSize});
  return $bytes;
}

sub readMiniSectorChain($)
{
  my $self = shift();
  my ($chain) = @_;

  my $bytes = '';
  for (my $i = 0; $i < scalar(@{$chain}); $i++)
  {
    # TODO: Combine reads?
    my $miniSector = $self->readMiniSector($chain->[$i]);
    (!defined($miniSector)) && return undef;
    $bytes .= $miniSector;
  }

  return $bytes;
}

# Reads one sector of the FAT.
sub readFatSector($)
{
  my $self = shift();
  my ($fatSectorIndex) = @_;

  my $s = $self->getDiFatEntry($fatSectorIndex);
  
  my $fat = $self->readSector($s);
  (!defined($fat)) && die("getFatEntry(): Couldn't read sector $s");

  return $fat;
}

sub getDiFatEntry($)
{
  my $self = shift();
  my ($index) = @_;

  ($index < 0) && die("getDiFatEntry(): Negative index");
  (!defined($self->{diFatMap})) && $self->readDiFatMap();
  (!defined($self->{diFatMap})) && return undef;

  if ($index < DIFAT_INLINE_SIZE)
  {
    return $self->{header}->{diFatInline}->[$index];
  }

  my $entriesPerSector = ($self->{sectorSize} >> 2) - 1;
  my $mapIndex = int(($index - 109) / $entriesPerSector);
  my $offset = (($index - 109) % $entriesPerSector);
  my $s = $self->{diFatMap}->[$mapIndex];
  (!defined($s)) && die("getDiFatEntry(): Requested index ${index} is outside of DiFat");

  my $sector = $self->readSector($s);
  (!defined($sector)) && die("getDiFatEntry(): Couldn't read sector $s");

  my $entry = unpack('V', substr($sector, $offset * 4, 4));
  return $entry;
}

sub getFatEntry($)
{
  my $self = shift();
  my ($index) = @_;

  ($index < 0) && die("getFatEntry(): Negative index");
  
  my $entriesPerSector = ($self->{sectorSize} >> 2);
  my $fatSectorIndex = int($index / $entriesPerSector);
  my $fatSectorOffset = $index % $entriesPerSector;

  my $fat = $self->readFatSector($fatSectorIndex);
  (!defined($fat)) && die("getFatEntry(): Couldn't read FAT sector at index $fatSectorIndex");
  
  my $entry = unpack('V', substr($fat, $fatSectorOffset * 4, 4));
  return $entry;
}

sub getFatChain($)
{
  my $self = shift();
  my ($start) = @_;

  my $chain = [];
  my $seenSector = {};
  my $nextSector = $start;
  while ($nextSector != SECTOR_END_OF_CHAIN)
  {
    defined($seenSector->{$nextSector}) && die("getFatChain(): Loop detected in sector chain");
    ($nextSector > SECTOR_MAX) && die("getFatChain(): Unexpected next sector number ${nextSector}");

    push(@{$chain}, $nextSector);
    $seenSector->{$nextSector} = 1;

    $nextSector = $self->getFatEntry($nextSector);
  }

  return $chain;
}

sub getMiniFatEntry($)
{
  my $self = shift();
  my ($index) = @_;

  ($index < 0) && die("getMiniFatEntry(): Negative index");

  my $entriesPerSector = $self->{sectorSize} / MINIFAT_ENTRY_SIZE;
  my $miniFatSectorIndex = int($index / $entriesPerSector);
  my $miniFatSectorOffset = $index % $entriesPerSector;

  my $miniFatChain = $self->getFatChain($self->{header}->{miniFatStartSector});
  ($miniFatSectorIndex >= scalar(@{$miniFatChain})) && die("getMiniFatEntry(): Index ${index} out of bounds");

  my $sector = $self->readSector($miniFatChain->[$miniFatSectorIndex]);

  my $bytes = substr($sector, $miniFatSectorOffset * MINIFAT_ENTRY_SIZE, MINIFAT_ENTRY_SIZE);
  my $decoder = StructDecoder->new($bytes);

  my $entry = $decoder->decodeField('u32');

  return $entry;
}

sub getMiniFatChain($)
{
  my $self = shift();
  my ($start) = @_;

  my $chain = [];
  my $seenMiniSector = {};
  my $nextMiniSector = $start;
  while ($nextMiniSector != SECTOR_END_OF_CHAIN)
  {
    defined($seenMiniSector->{$nextMiniSector}) && die("readMiniFatChain(): Loop detected in sector chain");
    ($nextMiniSector > SECTOR_MAX) && die("readMiniFatChain(): Unexpected next mini sector number ${nextMiniSector}");

    push(@{$chain}, $nextMiniSector);
    $seenMiniSector->{$nextMiniSector} = 1;

    $nextMiniSector = $self->getMiniFatEntry($nextMiniSector);
  }

  return $chain;
}

sub getDirectoryEntry($)
{
  my $self = shift();
  my ($index) = @_;

  ($index < 0) && die("getDirectoryIndex(): Negative index");
  ($index >= scalar(@{$self->{directory}})) && die("getDirectoryIndex(): Index out of bounds");

  return Spreadsheet::Nifty::CDF::DirectoryEntry->new($self, $self->{directory}->[$index]);
}

sub getRootDirectory()
{
  my $self = shift();

  return $self->getDirectoryEntry(0);
}

sub getObjectCount()
{
  my $self = shift();

  (defined($self->{objectCount})) && return $self->{objectCount};

  my $objectCount = 0;
  for my $e (@{$self->{directory}})
  {
    ($e->{type} == OBJ_TYPE_NONE) && next;
    $objectCount++;
  }

  $self->{objectCount} = $objectCount;
  return $objectCount;
}

sub getObjectBytes($)
{
  my $self = shift();
  my ($id) = @_;

  ($id == 0) && return undef;  # Entry 0's contents would be the mini stream itself

  my $dirEntry = $self->getDirectoryEntry($id);
  (!defined($dirEntry)) && return undef;
  ($dirEntry->type() == OBJ_TYPE_NONE) && return undef;  # Empty entry

  my $bytes;
  if ($dirEntry->size() < $self->{header}->{miniStreamMaxObjectSize})
  {
    my $miniChain = $self->getMiniFatChain($dirEntry->startSector());
    $bytes = $self->readMiniSectorChain($miniChain);
  }
  else
  {
    my $chain = $self->getFatChain($dirEntry->startSector());
    $bytes = $self->readSectorChain($chain);
  }

  $bytes = substr($bytes, 0, $dirEntry->{size});
  return $bytes;
}

#sub openStream($)
#{
#  my $self = shift();
#  my ($id) = @_;
#
#  ($id == 0) && return undef;  # Entry 0 is not a stream
#
#  my $dirEntry = $self->getDirectoryEntry($id);
#  (!defined($dirEntry)) && return undef;
#  ($dirEntry->type() == OBJ_TYPE_NONE) && return undef;  # Empty entry
#
#  my ($chainType, $chain);
#  if ($dirEntry->size() <= $self->{header}->{miniStreamMaxObjectSize})
#  {
#    $chainType = CHAIN_TYPE_MINI;
#    $chain = $self->getMiniFatChain($dirEntry->startSector());
#  }
#  else
#  {
#    $chainType = CHAIN_TYPE_NORMAL;
#    $chain = $self->getFatChain($dirEntry->startSector());
#  }
#
#  my $handle = Spreadsheet::Nifty::CDF::Handle->new($self, $chainType, $chain, $dirEntry->size());
#  return $handle;
#}

sub dumpMiniFat()
{
  my $self = shift();

  my $miniFatChain = $self->getFatChain($self->{header}->{miniFatStartSector});
  my $miniFat = $self->readSectorChain($miniFatChain);

  my $decoder = StructDecoder->new($miniFat);
  my $i = 0;
  while ($decoder->bytesLeft())
  {
    my $entry = $decoder->decodeField('u32');
    printf("miniFat entry %3d: %8X\n", $i++, $entry);
  }
}

1;
