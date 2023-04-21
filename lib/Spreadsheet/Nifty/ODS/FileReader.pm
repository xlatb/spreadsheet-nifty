#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::ODS::FileReader;

use Spreadsheet::Nifty::OpenDocument;
use Spreadsheet::Nifty::XMLReaderUtils;
use Spreadsheet::Nifty::ODS;
use Spreadsheet::Nifty::ODS::Sheet;
use Spreadsheet::Nifty::ODS::SheetReader;

use XML::LibXML qw(:libxml);
use XML::LibXML::Reader;
use Data::Dumper;

# === Class methods ===

sub new()
{
  my $class = shift();
  
  my $self = {};
  $self->{odf}      = Spreadsheet::Nifty::OpenDocument->new();
  $self->{manifest} = undef;
  $self->{workbook} = {};
  $self->{parked}   = [];
  $self->{debug}    = 0;
  
  bless($self, $class);

  return $self;
}

sub isFileSupported($)
{
  my $class = shift();
  my ($filename) = @_;

  my $odf = Spreadsheet::Nifty::OpenDocument->new();
  (!$odf->open($filename)) && return 0;

  my $mimetype = $odf->readMimetype();
  (scalar(grep({ $mimetype eq $_ } @{$Spreadsheet::Nifty::ODS::mimetypes})) == 0) && return 0;  # Unexpected MIME type

  return 1;
}

# === Instance methods ===

sub open($)
{
  my $self = shift();
  my ($filename) = @_;

  (!$self->{odf}->open($filename)) && return 0;

  return $self->read();
}

sub read()
{
  my $self = shift();

  # Read manifest
  ($self->{debug}) && printf("read manifest\n");
  $self->{manifest} = $self->{odf}->readManifest();

  # Read workbook
  ($self->{debug}) && printf("readWorkbook\n");
  $self->readWorkbook();

#  ($self->{debug}) && printf("readStyles\n");
#  $self->readStyles();

  ($self->{debug}) && printf("reading complete\n");
  return 1;
}

sub readStyles($)
{
  my $self = shift();

  ...;  
}

sub readWorkbook()
{
  my $self = shift();

  my $xmlns = $Spreadsheet::Nifty::ODS::namespaces->{office};

  my $zipReader = $self->{odf}->openMember('content.xml');
  (!$zipReader) && die("Couldn't open member 'content.xml'\n");

  my $xmlReader = XML::LibXML::Reader->new({IO => $zipReader});

  my $status = $xmlReader->read();
  ($status != 1) && return 0;

  # Check root element
  ($xmlReader->namespaceURI() ne $xmlns) && return 0;
  ($xmlReader->localName() ne 'document-content') && return 0;

  # Children of root should be: <office:scripts/>, <office:font-face-decls/>, <office:automatic-styles/>, <office:body/>
  while (($status = $xmlReader->read()) == 1)
  {
    #printf("readWorkbook() depth %d empty %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->isEmptyElement(), $xmlReader->nodeType(), $xmlReader->localName());

    ($xmlReader->nodeType() != XML_READER_TYPE_ELEMENT) && next;
    ($xmlReader->namespaceURI() ne $xmlns) && next;

    my $localName = $xmlReader->localName();
    if ($localName eq 'body')
    {
      # Should contain a single child: <office:spreadsheet/>
      (!Spreadsheet::Nifty::XMLReaderUtils->findChildElement($xmlReader, 'spreadsheet', $xmlns)) && die("expected <office:spreadsheet/> inside <office:body/>");

      $self->readSheetMetas($xmlReader);

      (!Spreadsheet::Nifty::XMLReaderUtils->atEndOfElement($xmlReader, 'spreadsheet', $xmlns)) && die("Misaligned after reading sheet metas");
      #printf("  depth %d empty %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->isEmptyElement(), $xmlReader->nodeType(), $xmlReader->localName());
    }
    else
    {
      #printf("  skipping kids...\n");
      Spreadsheet::Nifty::XMLReaderUtils->skipElement($xmlReader);  # Skip all children
    }
  }

  return 1;
}

sub readSheetMetas($)
{
  my $self = shift();
  my ($xmlReader) = @_;

  my $xmlns = $Spreadsheet::Nifty::ODS::namespaces->{table};

  ($xmlReader->localName() ne 'spreadsheet') && die("Expected <office:spreadsheet/> element");

  my $startDepth = $xmlReader->depth();
  my $sheets = [];

  if (Spreadsheet::Nifty::XMLReaderUtils->findChildElement($xmlReader, 'table', $xmlns))
  {
    do
    {
      #Spreadsheet::Nifty::XMLReaderUtils->dump($xmlReader, "readSheetMetas() START");
      push(@{$sheets}, $self->readSheetMeta($xmlReader));
      #Spreadsheet::Nifty::XMLReaderUtils->dump($xmlReader, "readSheetMetas() END  ");
    } while (Spreadsheet::Nifty::XMLReaderUtils->findSiblingElement($xmlReader, 'table', $xmlns));
  }

  $self->{workbook}->{sheets} = $sheets;
  #print Dumper($sheets);

  #$xmlReader->skipSiblings();

  my $endDepth = $xmlReader->depth();

  #printf("readSheetMetas(): depth %d empty %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->isEmptyElement(), $xmlReader->nodeType(), $xmlReader->localName());
  ($endDepth != $startDepth) && die("Expected depth ${startDepth} after reading sheet metas, got depth ${endDepth}");
}

sub readSheetMeta($)
{
  my $self = shift();
  my ($xmlReader) = @_;

  my $xmlns = $Spreadsheet::Nifty::ODS::namespaces->{table};

  ($xmlReader->localName() ne 'table') && die("Expected <table:table/> element");

  my $startDepth = $xmlReader->depth();

  #printf("readSheetMeta(): depth %d empty %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->isEmptyElement(), $xmlReader->nodeType(), $xmlReader->localName());
  
  my $meta = {};
  for my $k (qw(name protected protection-key protection-key-digest-algorithm))
  {
    $meta->{$k} = $xmlReader->getAttributeNs($k, $xmlns);
  }

  # Sadly, the number of used rows and columns is not stored anywhere in the
  #  file. Also sadly, the number of rows actually stored is often padded to
  #  some maximum value like 1048576. I don't think it's possible to reliably
  #  detect when we've reached the last non-empty row in a single pass.
  # Thus we're stuck having to scan through each sheet to collect this data.
  my $maxColCount = 0;
  my $maxRowCount = 0;
  if (Spreadsheet::Nifty::XMLReaderUtils->findChildElement($xmlReader, 'table-row', $xmlns))
  {
    my $rowIndex = 0;
    do
    {
      #Spreadsheet::Nifty::XMLReaderUtils->dump($xmlReader, "readSheetMeta() table-row");
      my $rowElement = $xmlReader->copyCurrentNode(1);
      Spreadsheet::Nifty::XMLReaderUtils->skipElement($xmlReader);

      # Decode rowDef (each can span multiple rows)
      my $rowDef = Spreadsheet::Nifty::ODS::Decode->decodeRowDefinition($rowElement);
      $rowDef->{startIndex} = $rowIndex;

      # Decode cellDefs (each can span multiple cells)
      my $cellDefs = [];
      my $children = [ $rowElement->getChildrenByTagName('*') ];
      for my $c (@{$children})
      {
        my $cellDef = Spreadsheet::Nifty::ODS::Decode->decodeCellDefinition($c);
        push(@{$cellDefs}, $cellDef);
      }

      # Discard empty trailing cellDefs
      while (scalar(@{$cellDefs}) && $cellDefs->[-1]->{empty})
      {
        pop(@{$cellDefs});
      }

      # Find remaining count of cells
      my $cellCount = 0;
      for my $c (@{$cellDefs})
      {
        $cellCount += $c->{count};
      }

      if ($cellCount > $maxColCount)
      {
        $maxColCount = $cellCount;
      }

      if ($cellCount > 0)
      {
        $maxRowCount = $rowIndex + $rowDef->{count};
      }

      $rowIndex += $rowDef->{count};

    } while (Spreadsheet::Nifty::XMLReaderUtils->findSiblingElement($xmlReader, 'table-row', $xmlns));
  }

  $meta->{dimensions}->{colCount} = $maxColCount;
  $meta->{dimensions}->{rowCount} = $maxRowCount;

  #Spreadsheet::Nifty::XMLReaderUtils->dump($xmlReader, "after dims:");

  Spreadsheet::Nifty::XMLReaderUtils->skipElement($xmlReader);

  return $meta;
}

sub parkXmlReaderForSheet($$)
{
  my $self = shift();
  my ($xmlReader, $sheetIndex) = @_;

  my $xmlns = $Spreadsheet::Nifty::ODS::namespaces->{table};

  $sheetIndex++;

  if ($sheetIndex >= scalar(@{$self->{workbook}->{sheets}}))
  {
    #printf("parkXmlReaderForSheet(): Final sheet; ignoring\n");
    return;  # It's from the final sheet so it's not worth reusing
  }
  elsif (!Spreadsheet::Nifty::XMLReaderUtils->atEndOfElement($xmlReader, 'table', $xmlns))
  {
    #printf("parkXmlReaderForSheet(): Not in expected position; ignoring\n");
    return;  # Reader not in expected position
  }

  # Advance to following table
  if (!Spreadsheet::Nifty::XMLReaderUtils->findSiblingElement($xmlReader, 'table', $xmlns))
  {
    #printf("parkXmlReaderForSheet(): No following table; ignoring\n");
    return;  # Can't find following table
  }

  #printf("Parked a reader for sheetIndex %d\n", $sheetIndex);
  push(@{$self->{parked}}, {sheetIndex => $sheetIndex, xmlReader => $xmlReader});
  return;
}

sub findParkedXmlReaderForSheet($)
{
  my $self = shift();
  my ($sheetIndex) = @_;

  # TODO: Also find an earlier reader and advance it to the requested sheet?

  my $readerIndex = undef;
  for (my $i = 0; $i < scalar(@{$self->{parked}}); $i++)
  {
    if ($sheetIndex == $self->{parked}->[$i]->{sheetIndex})
    {
      $readerIndex = $i;
      last;
    }
  }

  return $readerIndex;
}

sub xmlReaderForSheet($)
{
  my $self = shift();
  my ($sheetIndex) = @_;

  # Attempt to fulfill request with a parked reader
  my $parkIndex = $self->findParkedXmlReaderForSheet($sheetIndex);
  if (defined($parkIndex))
  {
    # Unpark
    #printf("Unparking reader for sheetIndex %d\n", $sheetIndex);
    my $item = splice(@{$self->{parked}}, $parkIndex, 1);
    return $item->{xmlReader};
  }

  my $zipReader = $self->{odf}->openMember('content.xml');
  (!$zipReader) && die("Couldn't open member 'content.xml'\n");

  my $xmlReader = XML::LibXML::Reader->new({IO => $zipReader});

  my $status = $xmlReader->read();
  ($status != 1) && return undef;

  # Sanity check root element
  (!Spreadsheet::Nifty::XMLReaderUtils->atStartOfElement($xmlReader, 'document-content', $Spreadsheet::Nifty::ODS::namespaces->{office})) && die("Unexpected root element");

  # Find child <office:body/>
  (!Spreadsheet::Nifty::XMLReaderUtils->findChildElement($xmlReader, 'body', $Spreadsheet::Nifty::ODS::namespaces->{office})) && return undef;

  # Find child <office:spreadsheet/>
  (!Spreadsheet::Nifty::XMLReaderUtils->findChildElement($xmlReader, 'spreadsheet', $Spreadsheet::Nifty::ODS::namespaces->{office})) && return undef;

  # Find first <table:table/>
  (!Spreadsheet::Nifty::XMLReaderUtils->findChildElement($xmlReader, 'table', $Spreadsheet::Nifty::ODS::namespaces->{table})) && return undef;
  ($sheetIndex == 0) && return $xmlReader;

  # Find subsequent sheets
  while (Spreadsheet::Nifty::XMLReaderUtils->findSiblingElement($xmlReader, 'table', $Spreadsheet::Nifty::ODS::namespaces->{table}))
  {
    $sheetIndex--;
    ($sheetIndex == 0) && return $xmlReader;
  }

  return undef;  # No such sheet
}

sub openSheet($)
{
  my $self = shift();
  my ($index) = @_;

  (($index < 0) || ($index >= scalar(@{$self->{workbook}->{sheets}}))) && return undef;  # Out of bounds

  ##my $xmlReader = $self->xmlReaderForSheet($index);
  ##(!defined($xmlReader)) && return undef;

  my $sheet = Spreadsheet::Nifty::ODS::Sheet->new($self, $index);
  (!$sheet->open()) && return undef;  # Failed to open

  return $sheet;
}

sub getSheetNames()
{
  my $self = shift();

  return [ map({ $_->{name} } @{$self->{workbook}->{sheets}}) ];
}

sub getSheetCount()
{
  my $self = shift();

  return scalar(@{$self->{workbook}->{sheets}});
}

sub getSheetRowCount($)
{
  my $self = shift();
  my ($sheetIndex) = @_;

  return $self->{workbook}->{sheets}->[$sheetIndex]->{dimensions}->{rowCount};
}

sub getSheetColCount($)
{
  my $self = shift();
  my ($sheetIndex) = @_;

  return $self->{workbook}->{sheets}->[$sheetIndex]->{dimensions}->{colCount};
}

1;
