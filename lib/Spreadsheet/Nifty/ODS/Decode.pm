#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::ODS::Decode;

use Spreadsheet::Nifty::ODS;

use XML::LibXML qw(:libxml);

sub decodeDateString($)
{
  my $class = shift();
  my ($str) = @_;

  # https://www.w3.org/TR/xmlschema-2/#dateTime
  if ($str =~ m#^(-?\d{4,})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(?:[.]\d+)?)(Z|[+-]\d{2}:\d{2})?$#)
  {
    my $year   = int($1);
    my $month  = int($2);
    my $day    = int($3);
    my $hour   = int($4);
    my $minute = int($5);
    my $second = int($6);
    my $tz     = $7;
  }
  # https://www.w3.org/TR/xmlschema-2/#date
  elsif ($str =~ m#^(-?\d{4,})-(\d{2})-(\d{2})(Z|[+-]\d{2}:\d{2})?$#)
  {
    my $year  = int($1);
    my $month = int($2);
    my $day   = int($3);
    my $tz    = $4;
  }

  return undef;
}

sub decodeTimeString($)
{
  my $class = shift();
  my ($str) = @_;

  # https://www.w3.org/TR/xmlschema-2/#duration
  if ($str =~ m#^(-?)P(\d+Y)?(\d+M)?(\d+D)?(?:T(\d+H)(\d+M)(\d+(?:[.]\d+)?S))?$#)
  {
    my $sign = defined($1) ? -1 : 1;
    my $years   = int($2 // 0);
    my $months  = int($3 // 0);
    my $days    = int($4 // 0);
    my $hours   = int($5 // 0);
    my $minutes = int($6 // 0);
    my $seconds = 0 + ($7 // 0);
  }  

  return undef;
}

sub collapseWhitespace($)
{
  my $class = shift();
  my ($str) = @_;

  $str =~ s/[\x09\x0D\x0A]/ /g;  # Non-space whitespace replaced by spaces
  $str =~ s/^ +//;  # Leading spaces removed
  $str =~ s/ +$//;  # Trailing spaces removed
  $str =~ s/  +/ /g;  # Runs of spaces replaced by single space

  return $str;
}

# Given an element, extracts its text content.
sub extractTextContent($)
{
  my $class = shift();
  my ($node) = @_;

  my $xmlns = $Spreadsheet::Nifty::ODS::namespaces->{text};

  my $text = '';

  my $children = [ $node->childNodes() ];
  for my $c (@{$children})
  {
    my $nodeType = $c->nodeType();
    if ($nodeType == XML_TEXT_NODE)
    {
      $text .= $class->collapseWhitespace($c->nodeValue());
    }
    elsif ($nodeType == XML_ELEMENT_NODE)
    {
      my $localname = $c->localname();
      my $ns = $c->namespaceURI();

      if ($ns eq $xmlns)
      {
        # <text:p>Regular text</text:p>
        # <text:h>Heading text</text:h>
        if (($localname eq 'p') || ($localname eq 'h'))
        {
          # These are block elements so we'll add a newline separator if not at the beginning of the text
          (length($text)) && do { $text .= "\n"; };
          $text .= $class->extractTextContent($c);
        }
        elsif (($localname eq 'span') || ($localname eq 'a') || ($localname eq 'number'))
        {
          $text .= $class->extractTextContent($c);
        }
        # <text:s />
        elsif ($localname eq 's')
        {
          my $count = int($c->getAttributeNS($xmlns, 'c') // 1);
          $text .= (' ' x $count);
        }
        # <text:tab />
        elsif ($localname eq 'tab')
        {
          $text .= "\x09";
        }
        elsif ($localname eq 'line-break')
        {
          $text .= "\n";
        }
        # TODO:
        # <text:list><text:list-header><text:h>Header</text:h></text:list-header><text:list-item><text:p>Item 1</text:p></text:list-item><text:list-item><text:p>Item 2</text:p></text:list-item></text:list>
      }
    }
  }

  return $text;
}

# <table:table-column table:style-name="co6" table:visibility="collapse" table:number-columns-repeated="238" table:default-cell-style-name="ce1" />
sub decodeColumnDefinition($)
{
  my $class = shift();
  my ($node) = @_;

  my $xmlns = $Spreadsheet::Nifty::ODS::namespaces->{table};

  my $columnDef = {};
  $columnDef->{count}      = int($node->getAttributeNS($xmlns, 'number-columns-repeated') // 1);
  $columnDef->{cellStyle}  = $node->getAttributeNS($xmlns, 'default-cell-style-name');
  $columnDef->{style}      = $node->getAttributeNS($xmlns, 'style-name');
  $columnDef->{visibility} = $node->getAttributeNS($xmlns, 'visibility') // 'visible';

  return $columnDef;
}

# <table:table-row table:style-name="ro3" table:number-rows-repeated="2">...</table:table-row>
sub decodeRowDefinition($)
{
  my $class = shift();
  my ($node) = @_;

  my $xmlns = $Spreadsheet::Nifty::ODS::namespaces->{table};

  my $rowDef = {};
  $rowDef->{count}      = int($node->getAttributeNS($xmlns, 'number-rows-repeated') // 1);
  $rowDef->{cellStyle}  = $node->getAttributeNS($xmlns, 'default-cell-style-name');
  $rowDef->{style}      = $node->getAttributeNS($xmlns, 'style-name');
  $rowDef->{visibility} = $node->getAttributeNS($xmlns, 'visibility') // 'visible';

  return $rowDef;
}

# A cell is either <table:table-cell/>, <table:covered-table-cell/>, or <table:change-track-table-cell/>
# <table:table-cell />
# <table:table-cell table:number-columns-repeated="1005" />
# <table:table-cell table:style-name="ce3" office:value-type="string" calcext:value-type="string"><text:p>Here is some text</text:p></table:table-cell>
# <table:table-cell table:style-name="ce101" office:value-type="float" office:value="1" calcext:value-type="float"><text:p>1.00</text:p></table:table-cell>
# <table:covered-table-cell table:number-columns-repeated="14" table:style-name="ce41" />
sub decodeCellDefinition($)
{
  my $class = shift();
  my ($node) = @_;

  my $tablens  = $Spreadsheet::Nifty::ODS::namespaces->{table};
  my $officens = $Spreadsheet::Nifty::ODS::namespaces->{office};

  my $valueType = $node->getAttributeNS($officens, 'value-type') // 'void';

  my $text;
  if ($valueType ne 'void')
  {
    $text = $class->extractTextContent($node);
  }

  my $value;
  if (($valueType eq 'currency') || ($valueType eq 'float') || ($valueType eq 'percentage'))
  {
    $value = $node->getAttributeNS($officens, 'value');
  }
  elsif ($valueType eq 'string')
  {
    $value = $node->getAttributeNS($officens, 'string-value');
    if (!defined($value))
    {
      $value = $text;
    }
  }
  elsif ($valueType eq 'boolean')
  {
    $value = $node->getAttributeNS($officens, 'boolean-value');
  }
  elsif ($valueType eq 'date')
  {
    $value = $node->getAttributeNS($officens, 'date-value');
  }
  elsif ($valueType eq 'time')
  {
    $value = $node->getAttributeNS($officens, 'time-value');
  }

  my $cellDef = {};
  $cellDef->{count}     = int($node->getAttributeNS($tablens, 'number-columns-repeated') // 1);
  $cellDef->{style}     = $node->getAttributeNS($tablens, 'style-name');
  $cellDef->{valueType} = $valueType;
  $cellDef->{value}     = $value;
  $cellDef->{text}      = $text;

  # If a cell has no interesting features, mark it as "empty".
  if (($valueType eq 'void') && !$node->hasChildNodes() && !defined($cellDef->{style}))
  {
    $cellDef->{empty} = !!1;
  }

  return $cellDef;
}

1;
