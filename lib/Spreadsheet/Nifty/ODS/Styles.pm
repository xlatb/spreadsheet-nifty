#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::ODS::Styles;

use XML::LibXML qw(:libxml);
use XML::LibXML::Reader;

# === Class methods ===

sub new()
{
  my $class = shift();

  my $self = {};
  $self->{family}        = {};  # Keyed by family and then name
  $self->{default}       = {};  # Keyed by family
  $self->{numberFormats} = {};  # Keyed by name
  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub read($)
{
  my $self = shift();
  my ($xmlReader) = @_;

  # Check current element
  ($xmlReader->namespaceURI() ne $Spreadsheet::Nifty::ODS::namespaces->{office}) && return !!0;
  ($xmlReader->localName() !~ m#^(styles|automatic-styles)$#) && return !!0;

  my $depth = $xmlReader->depth();

  my $styleNS = $Spreadsheet::Nifty::ODS::namespaces->{style};
  my $numberNS = $Spreadsheet::Nifty::ODS::namespaces->{number};

  # Expected children: <style:style/>, <style:default-style/>, <number:number-style/>, <number:date-style/>, <number:time-style/>, <number:currency-style/>, <number:percentage-style/>, <number:boolean-style/>
  while ($xmlReader->read() == 1)
  {
    #printf("Styles->read() depth %d empty %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->isEmptyElement(), $xmlReader->nodeType(), $xmlReader->localName());

    ($xmlReader->depth() == $depth) && ($xmlReader->nodeType() == XML_READER_TYPE_END_ELEMENT) && last;  # Found the end

    ($xmlReader->nodeType() != XML_READER_TYPE_ELEMENT) && next;

    my $ns = $xmlReader->namespaceURI();
    my $localName = $xmlReader->localName();

    if (($ns eq $styleNS) && ($localName eq 'style'))  # <style:style/>
    {
      my $family = $xmlReader->getAttributeNs('family', $styleNS);
      my $name   = $xmlReader->getAttributeNs('name', $styleNS);

      my $style = Spreadsheet::Nifty::ODS::Decode->decodeStyle($xmlReader->copyCurrentNode(1));
      if (defined($style))
      {
        $self->{family}->{$family}->{$name} = $style;
      }
    }
    elsif (($ns eq $styleNS) && ($localName eq 'default-style'))  # <style:default-style/>
    {
      my $family = $xmlReader->getAttributeNs('family', $styleNS);

      my $style = Spreadsheet::Nifty::ODS::Decode->decodeStyle($xmlReader->copyCurrentNode(1));
      if (defined($style))
      {
        $self->{default}->{$family} = $style;
      }
    }
    elsif ($ns eq $numberNS)
    {
      my $name = $xmlReader->getAttributeNs('name', $styleNS);

      if (($localName eq 'number-style') ||
          ($localName eq 'text-style') ||
          ($localName eq 'date-style') ||
          ($localName eq 'time-style') ||
          ($localName eq 'currency-style') ||
          ($localName eq 'percentage-style'))
      {
        $self->{numberFormats}->{$name} = Spreadsheet::Nifty::ODS::Decode->decodeNumberFormat($xmlReader->copyCurrentNode(1));
      }
    }

    Spreadsheet::Nifty::XMLReaderUtils->skipElement($xmlReader);  # Skip past current element
  }

  return !!1;
}

1;
