#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSX::FileReader;

use Spreadsheet::Nifty::XLSX;
use Spreadsheet::Nifty::XLSX::Decode;
use Spreadsheet::Nifty::XLSX::Sheet;
use Spreadsheet::Nifty::XLSX::SheetReader;
use Spreadsheet::Nifty::ZIPPackage;
use Spreadsheet::Nifty::IndexedColors;

use XML::LibXML qw(:libxml);
use XML::LibXML::Reader;
use Data::Dumper;

# === Class methods ===

sub new()
{
  my $class = shift();
  
  my $self = {};
  $self->{zipPackage}    = Spreadsheet::Nifty::ZIPPackage->new();
  $self->{debug}         = 0;
  $self->{workbook}      = undef;
  $self->{sharedStrings} = undef;
  $self->{relationships} = undef;
  $self->{filename}      = undef;
  
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

  ($rel->{target} !~ m#[.]xml$#i) && return 0;

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
  (!defined($self->{relationships})) && return 0;

  # Find workbook partname
  my $rel = $self->{zipPackage}->getRelationshipByType($self->{relationships}, $Spreadsheet::Nifty::ZIPPackage::namespaces->{officeDocument});
  (!defined($rel)) && return 0;
  my $partname = $rel->{partname};

  # Read workbook relationships
  $self->{workbook} = {};
  $self->{workbook}->{relationships} = $self->{zipPackage}->readRelationshipsMember($self->{zipPackage}->partnameToRelationshipsPart($partname));
  ($self->{debug}) && print main::Dumper($self->{workbook}->{relationships});

  # Read workbook
  $self->{debug} && printf("readWorkbook\n");
  $self->readWorkbook($partname);

  # Read shared strings
  $self->{debug} && printf("readSharedStrings\n");
  $self->readSharedStrings();
  
  # Read styles
  $self->{debug} && printf("readStyles\n");
  $self->readStyles();
  $self->{debug} && print main::Dumper($self->{numberFormats});

  $self->{debug} && printf("reading workbook complete\n");
  return 1;
}

# Given a hash of attributes for an fgColor or bgColor object, returns the RGB value.
# Possible colour types: auto, indexed, rgb, theme, tint
sub resolveColor($)
{
  my $self = shift();
  my ($c) = @_;

  # <fgColor rgb="FFFFFF00"/>
  if (defined($c->{rgb}))
  {
    # Microsoft uses ARGB order, so we normalize to RGBA
    if (length($c->{rgb}) == 8)
    {
      return substr($c->{rgb}, 2, 6) . substr($c->{rgb}, 0, 2);
    }
    elsif (length($c->{rgb}) == 4)
    {
      return substr($c->{rgb}, 1, 3) . substr($c->{rgb}, 0, 1);
    }

    return $c->{rgb};
  }

  # <fgColor indexed="64"/>
  if (defined($c->{indexed}))
  {
    return $self->{indexedColors}->getColor($c->{indexed});
  }

  # TODO: <fgColor theme="0"/>
  # TODO - Colours (of any type) may also have an optional 'tint' attribute to brighten or darken.

  return undef;
}

sub readSharedStrings()
{
  my $self = shift();
  
  # Find workbook's shared strings relationship
  my $rel = $self->{zipPackage}->getRelationshipByType($self->{workbook}->{relationships}, $Spreadsheet::Nifty::ZIPPackage::namespaces->{sharedStrings});
  (!defined($rel)) && return 0;

  my $xmlns = $Spreadsheet::Nifty::XLSX::namespaces->{main};
  
  my $zipReader = $self->{zipPackage}->openMember($rel->{partname});
  (!$zipReader) && return;
  
  my $xmlReader = XML::LibXML::Reader->new({IO => $zipReader});

  my $sharedStrings = [];
 
  my $status;
  while (($status = $xmlReader->nextElement('si', $xmlns)) == 1)
  {
    #printf("readSharedStrings() depth %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->nodeType(), $xmlReader->localName());
    my $node = $xmlReader->copyCurrentNode(1);
    
    my $str = Spreadsheet::Nifty::XLSX::Decode->decodePlainString($node);
    
    push(@{$sharedStrings}, $str);
  }

  $self->{sharedStrings} = $sharedStrings;
  
  return;
}

sub getSharedString($)
{
  my $self = shift();
  my ($i) = @_;

  (($i < 0) || ($i > scalar(@{$self->{sharedStrings}}))) && die("getSharedString(): Out of range");

  return $self->{sharedStrings}->[$i];
}

sub readStyles()
{
  my $self = shift();

  # Find workbook's styles relationship
  my $rel = $self->{zipPackage}->getRelationshipByType($self->{workbook}->{relationships}, $Spreadsheet::Nifty::ZIPPackage::namespaces->{styles});
  (!defined($rel)) && return 0;

  my $zipReader = $self->{zipPackage}->openMember($rel->{partname});
  (!$zipReader) && return 0;
  
  my $xmlns = $Spreadsheet::Nifty::XLSX::namespaces->{main};

  my $styles = [];
  my $fills = [];
  my $fonts = [];
  my $numberFormats = {};
  my $indexedColors = Spreadsheet::Nifty::IndexedColors->new();
  
  my $xmlReader = XML::LibXML::Reader->new({IO => $zipReader});
 
  my $status;
  while (($status = $xmlReader->nextElement()) == 1)
  {
    #printf("readStyles() depth %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->nodeType(), $xmlReader->localName());
    if (($xmlReader->localName() eq 'numFmts') && ($xmlReader->namespaceURI() eq $xmlns))
    {
      my $node = $xmlReader->copyCurrentNode(1);
      
      my $numFmts = [ $node->getChildrenByTagNameNS($xmlns, 'numFmt') ];
      for my $numFmt (@{$numFmts})
      {
        my $id = $numFmt->getAttribute('numFmtId');
        my $format = $numFmt->getAttribute('formatCode');
        
        $numberFormats->{$id} = $format;
      }
    }
    elsif (($xmlReader->localName() eq 'cellXfs') && ($xmlReader->namespaceURI() eq $xmlns))
    {
      my $node = $xmlReader->copyCurrentNode(1);
      
      my $xfs = [ $node->getChildrenByTagNameNS($xmlns, 'xf') ];
      for my $xf (@{$xfs})
      {
        my $numberFormatId = $xf->getAttribute('numFmtId');
        my $fillId = $xf->getAttribute('fillId');
        my $fontId = $xf->getAttribute('fontId');

        push(@{$styles}, {numberFormatId => $numberFormatId, fillId => $fillId, fontId => $fontId});
      }      
    }
    elsif (($xmlReader->localName() eq 'fills') && ($xmlReader->namespaceURI() eq $xmlns))
    {
      my $node = $xmlReader->copyCurrentNode(1);

      my $fillElements = [ $node->getChildrenByTagNameNS($xmlns, 'fill') ];
      for my $fillElement (@{$fillElements})
      {
        # Child is either <patternFill/> or <gradientFill/>
        my ($child) = $fillElement->getChildrenByTagNameNS($xmlns, '*');
        if ($child->localname() eq 'patternFill')
        {
          my $patternType = $child->getAttribute('patternType') // 'none';
          my ($bgcolor) = $child->getChildrenByTagNameNS($xmlns, 'bgColor');
          my ($fgcolor) = $child->getChildrenByTagNameNS($xmlns, 'fgColor');

          # NOTE: FG/BG swapped for solid fill
          if ($patternType eq 'solid')
          {
            ($fgcolor, $bgcolor) = ($bgcolor, $fgcolor);
          }

          my $fill = {type => 'pattern', patternType => $patternType, fgColor => Spreadsheet::Nifty::XLSX::Decode->attributesHash($fgcolor), bgColor => Spreadsheet::Nifty::XLSX::Decode->attributesHash($bgcolor)};
          #printf("Fill element: %s\n", $fillElement->toString());
          #printf("Fill object: %s\n", main::Dumper($fill));
          push(@{$fills}, $fill)
        }
        elsif ($child->localname() eq 'gradientFill')
        {
          push(@{$fills}, {type => 'gradient'});
        }
        else
        {
          push(@{$fills}, {type => 'unknown'});
        }
      }
    }
    elsif (($xmlReader->localName() eq 'colors') && ($xmlReader->namespaceURI() eq $xmlns))
    {
      my $node = $xmlReader->copyCurrentNode(1);

      my ($indexedColorsElement) = $node->getChildrenByTagNameNS($xmlns, 'indexedColors');
      if (defined($indexedColorsElement))
      {
        my $rgbColors = [ $indexedColorsElement->getChildrenByTagNameNS($xmlns, 'rgbColor') ];
        for my $rgbColor (@{$rgbColors})
        {
          $indexedColors->addColor($rgbColor->getAttribute('rgb'));
        }
      }
    }
    elsif (($xmlReader->localName() eq 'fonts') && ($xmlReader->namespaceURI() eq $xmlns))
    {
      my $node = $xmlReader->copyCurrentNode(1);

      # TODO: <sz val="9"/>
      # TODO: <name val="Segoe UI"/>

      my $fontElements = [ $node->getChildrenByTagNameNS($xmlns, 'font') ];
      for my $fontElement (@{$fontElements})
      {
        #printf("FONT: %s\n", $fontElement->toString());
        my $fontdata = {};

        my ($color) = $fontElement->getChildrenByTagNameNS($xmlns, 'color');
        if (defined($color))
        {
          my $attrs = Spreadsheet::Nifty::XLSX::Decode->attributesHash($color);
          $fontdata->{color} = $attrs;
        }

        #printf("  FONTDATA: %s\n", main::Dumper($fontdata));

        push(@{$fonts}, $fontdata);
      }

    }
  }

  $self->{styles} = $styles;
  $self->{fills} = $fills;
  $self->{numberFormats} = $numberFormats;
  $self->{indexedColors} = $indexedColors;
  $self->{fonts} = $fonts;

  #print main::Dumper('styles', $self->{styles});
  #print main::Dumper('fonts', $self->{fonts});
  #print main::Dumper('indexedColors', $self->{indexedColors});

  return;
}

sub readWorkbook($)
{
  my $self = shift();
  my ($partname) = @_;
  
  my $xmlns = $Spreadsheet::Nifty::XLSX::namespaces->{main};
  my $relns = $Spreadsheet::Nifty::ZIPPackage::namespaces->{relationshipRefs};

  # Read workbook data member
  my $zipReader = $self->{zipPackage}->openMember($partname);
  ($zipReader) || die("Can't open workbook file '${partname}'\n");

  my $xmlReader = XML::LibXML::Reader->new({IO => $zipReader});
 
  my $status;
  while (($status = $xmlReader->nextElement()) == 1)
  {
   #printf("readWorkbook() depth %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->nodeType(), $xmlReader->localName());
   ($xmlReader->namespaceURI() eq $xmlns) || next;
    
    if ($xmlReader->localName() eq 'workbookPr')
    {
      $self->{workbook}->{flagDate1904} = Spreadsheet::Nifty::XLSX::Decode->decodeBool($xmlReader->getAttribute('date1904')) // !!0;
    }
    elsif ($xmlReader->localName() eq 'sheets')
    {
      my $node = $xmlReader->copyCurrentNode(1);

      my $kids = [ $node->getChildrenByTagNameNS($xmlns, 'sheet') ];

      my $sheets = [];
      for my $kid (@{$kids})
      {
        my $sheet = {};
        $sheet->{name} = $kid->getAttribute('name');
        $sheet->{id}   = $kid->getAttributeNS($relns, 'id');
        
        push(@{$sheets}, $sheet);
      }

      $self->{workbook}->{sheets} = $sheets;
    }
    elsif ($xmlReader->localName() eq 'workbookProtection')
    {
      my $node = $xmlReader->copyCurrentNode(1);

      $self->{workbook}->{protection} = Spreadsheet::Nifty::XLSX::Decode->decodeWorkbookProtection($node);
    }
  }

  return;
}

sub openSheet($)
{
  my $self = shift();
  my ($index) = @_;

  (($index < 0) || ($index >= scalar(@{$self->{workbook}->{sheets}}))) && return undef;  # Out of bounds

  # Find worksheet's relationship
  my $relId = $self->{workbook}->{sheets}->[$index]->{id};
  my $rel = $self->{workbook}->{relationships}->{$relId};
  (!defined($rel)) && return undef;

  my $zipReader = $self->{zipPackage}->openMember($rel->{partname});
  (!$zipReader) && return undef;

  my $reader = Spreadsheet::Nifty::XLSX::SheetReader->new($self, $index, $zipReader);
  (!$reader->open()) && return undef;

  my $sheet = Spreadsheet::Nifty::XLSX::Sheet->new($reader);
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

1;
