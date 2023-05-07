#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSB::FileReader;

use Spreadsheet::Nifty::XLSB;
use Spreadsheet::Nifty::XLSB::RecordReader;
use Spreadsheet::Nifty::XLSB::RecordTypes;
use Spreadsheet::Nifty::XLSB::SheetReader;
use Spreadsheet::Nifty::XLSB::Sheet;
use Spreadsheet::Nifty::XLSB::Decode;
use Spreadsheet::Nifty::StructDecoder;
use Spreadsheet::Nifty::ZIPPackage;

use Data::Dumper;

# === Class methods ===

sub new()
{
  my $class = shift();
  
  my $self = {};
  $self->{filename}      = undef;
  $self->{zipPackage}    = Spreadsheet::Nifty::ZIPPackage->new();
  $self->{members}       = {};
  $self->{debug}         = 0;
  $self->{workbook}      = undef;
  $self->{relationships} = undef;
  $self->{sharedStrings} = undef;
#  $self->{worksheets}    = undef;
  
  bless($self, $class);

  return $self;
}

sub isFileSupported($)
{
  my $class = shift();
  my ($filename) = @_;

  my $zippkg = Spreadsheet::Nifty::ZIPPackage->new();
  (!$zippkg->open($filename)) && return 0;

  my $rels = $zippkg->readRelationshipsMember('_rels/.rels');
  (!defined($rels)) && return 0;

  my $rel = $zippkg->getRelationshipByType($rels, $Spreadsheet::Nifty::ZIPPackage::namespaces->{officeDocument});
  (!defined($rel)) && return 0;

  ($rel->{target} !~ m#[.]bin$#i) && return 0;

  return 1;
}

# === Instance methods ===

sub open($)
{
  my $self = shift();
  my ($filename) = @_;

  (!$self->{zipPackage}->open($filename)) && return 0;

  return $self->read();
}

sub read()
{
  my $self = shift();

  # Read top-level relationships
  ($self->{debug}) && printf("read relationships\n");
  $self->{relationships} = $self->{zipPackage}->readRelationshipsMember('_rels/.rels');

  # Read workbook
  ($self->{debug}) && printf("readWorkbook\n");
  $self->readWorkbook();

  ($self->{debug}) && printf("readSharedStrings\n");
  $self->readSharedStrings();
  
  #($self->{debug}) && printf("readStyles\n");
  #$self->readStyles();
  #($self->{debug}) && print main::Dumper($self->{numberFormats});

  #($self->{debug}) && printf("readRelationships\n");
  #$self->readRelationships();
  
  #($self->{debug}) && printf("readWorkbook\n");
  #$self->readWorkbook();
  
  #for my $sheet (@{$self->{worksheets}})
  #{
  #  ($self->{debug}) && printf("readWorksheet\n");
  #  $self->readWorksheet($sheet);
  #}
  
  ($self->{debug}) && printf("reading complete\n");
  return 1;
}

sub readSharedStrings()
{
  my $self = shift();

  # Find workbook's shared strings relationship
  my $rel = $self->{zipPackage}->getRelationshipByType($self->{workbook}->{relationships}, $Spreadsheet::Nifty::ZIPPackage::namespaces->{sharedStrings});
  (!defined($rel)) && return 0;

  $self->{sharedStrings} = [];

  my $reader = $self->{zipPackage}->openMember($rel->{partname});
  (!$reader) && die("Couldn't open member '$rel->{partname}'\n");

  my $recs = Spreadsheet::Nifty::XLSB::RecordReader->new($reader);
  while (my $rec = $recs->read())
  {
    #printf("REC type %d size %d name %s\n", $rec->{type}, $rec->{size}, $rec->{name} // '?');
    #($rec->{size}) && printf("  data: %s\n", unpack('H*', $rec->{data}));

    if ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::SST_ITEM)
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $item = $decoder->decodeHash(['flags:u8', 'string:XLWideString']);  # TODO: Following fields
      push(@{$self->{sharedStrings}}, $item);
    }
  }
  
  return;
}

sub readWorkbook()
{
  my $self = shift();

  $self->{workbook} = {};
  
  #print main::Dumper($self->{relationships});
  my $rel = $self->{zipPackage}->getRelationshipByType($self->{relationships}, $Spreadsheet::Nifty::ZIPPackage::namespaces->{officeDocument});
  (!defined($rel)) && return 0;

  $self->{workbook}->{partname} = $rel->{partname};

  my $reader = $self->{zipPackage}->openMember($self->{workbook}->{partname});
  (!$reader) && die("Couldn't open member '$self->{workbook}->{partname}'\n");

  my $recs = Spreadsheet::Nifty::XLSB::RecordReader->new($reader);
  while (my $rec = $recs->read())
  {
    ($self->{debug}) && printf("REC type %d size %d name %s\n", $rec->{type}, $rec->{size}, $rec->{name} // '?');
    ($self->{debug}) && ($rec->{size}) && printf("  data: %s\n", unpack('H*', $rec->{data}));

    if ($rec->{name} eq 'BrtAbsPath15')
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $path = $decoder->decodeField('XLWideString');  # Original file path
    }
    elsif ($rec->{name} eq 'BrtBundleSh')
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $sheet = $decoder->decodeHash(['hsState:u32', 'iTabID:u32', 'strRelID:XLNullableWideString', 'strName:XLWideString']);
      (!defined($self->{workbook}->{sheets})) && do { $self->{workbook}->{sheets} = []; };
      push(@{$self->{workbook}->{sheets}}, $sheet);
      ($self->{debug}) && print main::Dumper($sheet);
    }
  }

  # Read workbook relationships
  $self->{workbook}->{relationships} = $self->{zipPackage}->readRelationshipsMember($self->{zipPackage}->partnameToRelationshipsPart($self->{workbook}->{partname}));
  ($self->{debug}) && print main::Dumper($self->{workbook}->{relationships});

  return 1;
}

sub openSheet($)
{
  my $self = shift();
  my ($index) = @_;

  (($index < 0) || ($index >= scalar(@{$self->{workbook}->{sheets}}))) && return undef;  # Out of bounds

  # Find workbook's shared strings relationship
  my $relId = $self->{workbook}->{sheets}->[$index]->{strRelID};
  my $rel = $self->{workbook}->{relationships}->{$relId};
  (!defined($rel)) && return undef;

  my $reader = Spreadsheet::Nifty::XLSB::SheetReader->new($self, $index, $rel->{partname});
  $reader->open();

  my $sheet = Spreadsheet::Nifty::XLSB::Sheet->new($self, $reader);
  return $sheet;
}

sub getSheetNames()
{
  my $self = shift();

  return [ map({ $_->{strName} } @{$self->{workbook}->{sheets}}) ];
}

sub getSheetCount()
{
  my $self = shift();

  return scalar(@{$self->{workbook}->{sheets}});
}

1;
