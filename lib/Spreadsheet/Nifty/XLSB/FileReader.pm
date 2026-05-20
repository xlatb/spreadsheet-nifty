#!/usr/bin/perl -w
use warnings;
use strict;

use Data::Dumper;

package Spreadsheet::Nifty::XLSB::FileReader;

use Spreadsheet::Nifty::XLSB;
use Spreadsheet::Nifty::XLSB::RecordReader;
use Spreadsheet::Nifty::XLSB::RecordTypes;
use Spreadsheet::Nifty::XLSB::SheetReader;
use Spreadsheet::Nifty::XLSB::Sheet;
use Spreadsheet::Nifty::XLSB::Styles;
use Spreadsheet::Nifty::XLSB::Decode;
use Spreadsheet::Nifty::StructDecoder;
use Spreadsheet::Nifty::ZIPPackage;

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
  $self->{styles}        = undef;
  $self->{externals}     = {links => [], sheets => []};
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
  ($self->{debug}) && printf("FileReader read relationships\n");
  $self->{relationships} = $self->{zipPackage}->readRelationshipsMember('_rels/.rels');

  # Read workbook
  ($self->{debug}) && printf("FileReader readWorkbook\n");
  $self->readWorkbook();

  ($self->{debug}) && printf("FileReader readSharedStrings\n");
  $self->readSharedStrings();

  ($self->{debug}) && printf("FileReader readStyles\n");
  $self->readStyles();
  
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

sub readStyles()
{
  my $self = shift();

  # Find workbook's styles relationship
  my $rel = $self->{zipPackage}->getRelationshipByType($self->{workbook}->{relationships}, $Spreadsheet::Nifty::ZIPPackage::namespaces->{styles});
  (!defined($rel)) && return 0;

  my $reader = $self->{zipPackage}->openMember($rel->{partname});
  (!$reader) && die("Couldn't open member '$rel->{partname}'\n");

  $self->{styles} = Spreadsheet::Nifty::XLSB::Styles->new();
  $self->{styles}->read($reader);

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
  while (defined(my $rec = $recs->read()))
  {
    ($self->{debug}) && printf("readWorkbook(): REC type %d size %d name %s\n", $rec->{type}, $rec->{size}, $rec->{name} // '?');
    ($self->{debug}) && ($rec->{size}) && printf("  data: %s\n", unpack('H*', $rec->{data}));

    if ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::ABS_PATH_15)
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $path = $decoder->decodeField('XLWideString');  # Original file path
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::BUNDLE_SH)
    {
      # TODO: Check for record BEGIN_BUNDLE_SHS and then read contents
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $sheet = $decoder->decodeHash(['hsState:u32', 'iTabID:u32', 'strRelID:XLNullableWideString', 'strName:XLWideString']);
      (!defined($self->{workbook}->{sheets})) && do { $self->{workbook}->{sheets} = []; };
      push(@{$self->{workbook}->{sheets}}, $sheet);
      ($self->{debug}) && print main::Dumper($sheet);
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::BEGIN_EXTERNALS)
    {
      $self->readWorkbookExternals($recs);
    }
  }

  # Read workbook relationships
  $self->{workbook}->{relationships} = $self->{zipPackage}->readRelationshipsMember($self->{zipPackage}->partnameToRelationshipsPart($self->{workbook}->{partname}));
  ($self->{debug}) && print main::Dumper($self->{workbook}->{relationships});

  return 1;
}

sub readWorkbookExternals($)
{
  my $self = shift();
  my ($recs) = @_;

  $self->{externals}->{links} = [];
  $self->{externals}->{sheets} = [];

  while (defined(my $rec = $recs->read()))
  {
    ($self->{debug}) && printf("readWorkbookExternals(): REC type %d size %d name %s\n", $rec->{type}, $rec->{size}, $rec->{name} // '?');

    if ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::END_EXTERNALS)
    {
      return !!1;  # All done
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::SUP_SELF)
    {
      my $link = {type => 'workbook-self'};
      push(@{$self->{externals}->{links}}, $link);
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::SUP_SAME)
    {
      my $link = {type => 'sheet-self'};
      push(@{$self->{externals}->{links}}, $link);
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::SUP_BOOK_SRC)
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $link = {type => 'external'};
      $link->{relId} = $decoder->decodeField('relID:XLNullableWideString');
      push(@{$self->{externals}->{links}}, $link);
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::SUP_ADDIN)
    {
      my $link = {type => 'addin', rec => $rec};
      push(@{$self->{externals}->{links}}, $link);
    }
    elsif ($rec->{type} == Spreadsheet::Nifty::XLSB::RecordTypes::EXTERN_SHEET)
    {
      my $decoder = Spreadsheet::Nifty::XLSB::Decode->decoder($rec->{data});
      my $count = $decoder->decodeField('u32');
      for (my $i = 0; $i < $count; $i++)
      {
        my $sheet = $decoder->decodeHash(['linkId:u32', 'minSheet:i32', 'maxSheet:i32']);
        push(@{$self->{externals}->{sheets}}, $sheet);
      }
    }
  }

  return !!0;  # EOF before finding BrtEndExternals
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

sub getExternalSheetReference($)
{
  my $self = shift();
  my ($index) = @_;

  (($index < 0) || ($index >= scalar(@{$self->{externals}->{sheets}}))) && return undef;  # External sheet out of bounds
  my $externalSheet = $self->{externals}->{sheets}->[$index];

  my $linkId = $externalSheet->{linkId};
  (($linkId < 0) || ($linkId >= scalar(@{$self->{externals}->{links}}))) && return undef;  # External link out of bounds
  my $externalLink = $self->{externals}->{links}->[$linkId];

  return {externalSheet => $externalSheet, externalLink => $externalLink};
}

1;
