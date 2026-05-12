#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::ODS::Decode;

use Spreadsheet::Nifty::ODS;

use XML::LibXML qw(:libxml);

sub gatherAttributes($$$$)
{
  my $class = shift();
  my ($target, $element, $ns, $names) = @_;

  my $count = 0;

  for my $name (@{$names})
  {
    if ($element->hasAttributeNS($ns, $name))
    {
      $target->{$name} = $element->getAttributeNS($ns, $name);
      $count++;
    }
  }

  return $count;
}

sub decodeBoolean($)
{
  my $class = shift();
  my ($str) = @_;

  # https://www.w3.org/TR/xmlschema-2/#boolean
  if (($str eq 'true') || ($str eq '1'))
  {
    return !!1;
  }
  elsif (($str eq 'false') || ($str eq '0'))
  {
    return !!0;
  }

  return undef;
}

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
    return {year => $year, month => $month, day => $day, hour => $hour, minute => $minute, second => $second, tz => $tz};
  }
  # https://www.w3.org/TR/xmlschema-2/#date
  elsif ($str =~ m#^(-?\d{4,})-(\d{2})-(\d{2})(Z|[+-]\d{2}:\d{2})?$#)
  {
    my $year  = int($1);
    my $month = int($2);
    my $day   = int($3);
    my $tz    = $4;
    return {year => $year, month => $month, day => $day, tz => $tz};
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

  my $formula = $node->getAttributeNS($tablens, 'formula');
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
  $cellDef->{f}         = $formula;

  # If a cell has no interesting features, mark it as "empty".
  if (($valueType eq 'void') && !$node->hasChildNodes() && !defined($cellDef->{style}))
  {
    $cellDef->{empty} = !!1;
  }

  return $cellDef;
}

# Handles either <style:style/> or <style:default-style/>
sub decodeStyle($)
{
  my $class = shift();
  my ($node) = @_;

  my $styleNS = $Spreadsheet::Nifty::ODS::namespaces->{style};
  my $family = $node->getAttributeNS($styleNS, 'family');

  my $style;
  if ($family eq 'table-cell')
  {
    $style = Spreadsheet::Nifty::ODS::Decode->decodeCellStyle($node);
  }
  elsif ($family eq 'table-column')
  {
    $style = Spreadsheet::Nifty::ODS::Decode->decodeColumnStyle($node);
  }
  elsif ($family eq 'table-row')
  {
    $style = Spreadsheet::Nifty::ODS::Decode->decodeRowStyle($node);
  }
  elsif ($family eq 'table')
  {
    $style = Spreadsheet::Nifty::ODS::Decode->decodeTableStyle($node);
  }
  else
  {
    return undef;
  }

  if ($node->hasAttributeNS($styleNS, 'parent-style-name'))
  {
    $style->{'parent-style-name'} = $node->getAttributeNS($styleNS, 'parent-style-name');
  }

  return $style;
}

# <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="xyz">
#   <style:table-cell-properties fo:background-color="#ffffff" fo:border="0.5pt solid #000000" style:vertical-align="middle" />
#   <style:text-properties fo:color="#333333" fo:font-weight="bold" style:font-name="arial" fo:font-size="15pt" />
# </style:style>
sub decodeCellStyle($)
{
  my $class = shift();
  my ($node) = @_;

  my $styleNS = $Spreadsheet::Nifty::ODS::namespaces->{style};
  my $foNS = $Spreadsheet::Nifty::ODS::namespaces->{fo};

  my $style = {};
  $style->{name} = $node->getAttributeNS($styleNS, 'name');

  my ($cell) = $node->getElementsByTagNameNS($styleNS, 'table-cell-properties');
  if (defined($cell))
  {
    $style->{cell} = {};
    $class->gatherAttributes($style->{cell}, $cell, $foNS, ['background-color', 'border', 'border-left', 'border-right', 'border-top', 'border-bottom', 'cell-protect']);
    $class->gatherAttributes($style->{cell}, $cell, $styleNS, ['cell-protect']);
  }

  my ($text) = $node->getElementsByTagNameNS($styleNS, 'text-properties');
  if (defined($text))
  {
    $style->{text} = {};
    $class->gatherAttributes($style->{text}, $text, $foNS, ['color', 'font-weight', 'font-name', 'font-size']);
  }

  return $style;
}

# <style:style style:name="co1" style:family="table-column">
#   <style:table-column-properties fo:break-before="auto" style:column-width="5.29mm" />
# </style:style>
sub decodeColumnStyle($)
{
  my $class = shift();
  my ($node) = @_;

  my $styleNS = $Spreadsheet::Nifty::ODS::namespaces->{style};

  my $style = {};
  $style->{name} = $node->getAttributeNS($styleNS, 'name');

  my ($column) = $node->getElementsByTagNameNS($styleNS, 'table-column-properties');
  if (defined($column))
  {
    $style->{column} = {};
    $class->gatherAttributes($style->{column}, $column, $styleNS, ['column-width']);
  }

  return $style;
}

sub decodeRowStyle($)
{
  my $class = shift();
  my ($node) = @_;

  my $styleNS = $Spreadsheet::Nifty::ODS::namespaces->{style};
  my $foNS = $Spreadsheet::Nifty::ODS::namespaces->{fo};

  my $style = {};
  $style->{name} = $node->getAttributeNS($styleNS, 'name');

  my ($row) = $node->getElementsByTagNameNS($styleNS, 'table-row-properties');
  if (defined($row))
  {
    $style->{row} = {};
    $class->gatherAttributes($style->{row}, $row, $styleNS, ['row-height']);
    $class->gatherAttributes($style->{row}, $row, $foNS, ['background-color']);
  }

  return $style;
}

sub decodeTableStyle($)
{
  my $class = shift();
  my ($node) = @_;

  my $styleNS = $Spreadsheet::Nifty::ODS::namespaces->{style};

  my $style = {};
  $style->{name} = $node->getAttributeNS($styleNS, 'name');

  my ($table) = $node->getElementsByTagNameNS($styleNS, 'table-properties');
  if (defined($table))
  {
    $style->{table} = {};
    $class->gatherAttributes($style->{table}, $table, $Spreadsheet::Nifty::ODS::namespaces->{tableooo}, ['tab-color']);
    $class->gatherAttributes($style->{table}, $table, $Spreadsheet::Nifty::ODS::namespaces->{loext}, ['tab-color']);
    $class->gatherAttributes($style->{table}, $table, $Spreadsheet::Nifty::ODS::namespaces->{table}, ['display', 'tab-color']);
    $class->gatherAttributes($style->{table}, $table, $Spreadsheet::Nifty::ODS::namespaces->{fo}, ['background-color']);
  }

  return $style;
}

# Decodes any of the following:
# • number:number-style:
#   <number:number-style style:name="N1">
#     <number:number number:decimal-places="0" loext:min-decimal-places="0" number:min-integer-digits="1"/>
#   </number:number-style>
# • number:currency-style
#   <number:currency-style style:name="N1">
#     <number:currency-symbol number:language="en" number:country="CA">$</number:currency-symbol>
#     <number:number number:decimal-places="2" loext:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/>
#   </number:currency-style>
# • number:percentage-style
#   <number:percentage-style style:name="N1">
#     <number:number number:decimal-places="2" loext:min-decimal-places="2" number:min-integer-digits="1"/>
#     <number:text>%</number:text>
#   </number:percentage-style>
# • number:date-style:
#   <number:date-style style:name="N1" number:language="en" number:country="CA">
#     <number:day number:style="long"/>
#     <number:text>-</number:text>
#     <number:month number:textual="true"/>
#     <number:text>-</number:text>
#     <number:year/>
#   </number:date-style>
# • number:time-style:
#   <number:time-style style:name="N1">
#     <number:minutes number:style="long"/>
#     <number:text>:</number:text>
#     <number:seconds number:style="long"/>
#   </number:time-style>
# • number:text-style:
#   <number:text-style style:name="N1">
#     <number:text-content/>
#   </number:text-style>
sub decodeNumberFormat()
{
  my $class = shift();
  my ($node) = @_;

  my $styleNS = $Spreadsheet::Nifty::ODS::namespaces->{style};
  my $numberNS = $Spreadsheet::Nifty::ODS::namespaces->{number};
  my $foNS = $Spreadsheet::Nifty::ODS::namespaces->{fo};

  my $style = {};
  $style->{name} = $node->getAttributeNS($styleNS, 'name');
  $style->{type} = $node->localname();
  $style->{items} = [];

  # Gather all children within the 'number' namespace
  my $children = [ $node->childNodes() ];
  for my $c (@{$children})
  {
    ($c->nodeType() != XML_ELEMENT_NODE) && next;  # Filter out comments or whitespace
    ($c->namespaceURI() ne $numberNS) && next;

    my $type = $c->localName();
    my $namespace = $c->namespaceURI();

    my $item = {};
    $item->{type} = $type;

    if ($namespace eq $numberNS)
    {
      # TODO: Some of these can have further <number:embedded-text/> child elements
	    # NOTE: The following are also valid, but have no attributes or children:
	    # • <number:text-content/>
	    # • <number:am-pm/>
	    # • <number:boolean/>
	    if ($type eq 'number')
	    {
	      $class->gatherAttributes($item, $c, $Spreadsheet::Nifty::ODS::namespaces->{loext}, ['min-decimal-places']);
	      $class->gatherAttributes($item, $c, $numberNS, ['decimal-places', 'min-integer-digits', 'min-decimal-places', 'grouping']);
	    }
	    elsif ($type eq 'scientific-number')
	    {
	      $class->gatherAttributes($item, $c, $Spreadsheet::Nifty::ODS::namespaces->{loext}, ['min-decimal-places']);
	      $class->gatherAttributes($item, $c, $numberNS, ['decimal-places', 'min-decimal-places', 'min-integer-digits', 'min-exponent-digits', 'forced-exponent-sign', 'grouping']);
	    }
	    elsif ($type eq 'fraction')
	    {
	      $class->gatherAttributes($item, $c, $Spreadsheet::Nifty::ODS::namespaces->{loext}, ['max-denominator-value']);
	      $class->gatherAttributes($item, $c, $numberNS, ['min-numerator-digits', 'min-denominator-digits', 'denominator-value', 'grouping']);
	    }
	    elsif ($type eq 'currency-symbol')
	    {
	      $class->gatherAttributes($item, $c, $numberNS, ['language', 'country']);
	      $item->{content} = $c->textContent();
	    }
	    elsif (($type eq 'day') || ($type eq 'year') || ($type eq 'era') || ($type eq 'day-of-week') || ($type eq 'quarter'))
	    {
	      $class->gatherAttributes($item, $c, $numberNS, ['calendar', 'style']);
	    }
	    elsif ($type eq 'month')
	    {
	      $class->gatherAttributes($item, $c, $numberNS, ['calendar', 'style', 'possessive-form', 'textual']);
	    }
	    elsif ($type eq 'week-of-year')
	    {
	      $class->gatherAttributes($item, $c, $numberNS, ['calendar']);
	    }
	    elsif (($type eq 'hours') || ($type eq 'minutes'))
	    {
	      $class->gatherAttributes($item, $c, $numberNS, ['style']);
	    }
	    elsif ($type eq 'seconds')
	    {
	      $class->gatherAttributes($item, $c, $numberNS, ['decimal-places', 'style']);
	    }
	    elsif ($type eq 'fill-character')
	    {
	      $item->{content} = $c->textContent();
	    }
	    elsif ($type eq 'text')
	    {
	      $item->{content} = $c->textContent();
	    }
      elsif ($type eq 'color')
      {
        # Ancient method of embedding a colour that predates ODF 1.0. We
        #  pretend it was a <style:text-propreties/>.
        $item->{type} = 'text-properties';
        $class->gatherAttributes($item, $c, $numberNS, ['color']);
      }
      push(@{$style->{items}}, $item);
    }
    elsif ($namespace eq $styleNS)
    {
      if ($type eq 'text-properties')
      {
        $class->gatherAttributes($item, $c, $foNS, ['color', 'background-color',
                                                    'country', 'language', 'script',
                                                    'font-family', 'font-size', 'font-style', 'font-variant', 'font-weight', 'font-name',
                                                    'text-underline-color', 'text-underline-mode', 'text-underline-style', 'text-underline-type', 'text-underline-width']);
        push(@{$style->{items}}, $item);
      }
      elsif ($type eq 'map')
      {
        # This will be stored under a separate 'maps' key, not as an 'item'.
        (!defined($style->{maps})) && do { $style->{maps} = []; };
        my $map = {};
        $class->gatherAttributes($map, $c, $styleNS, ['apply-style-name', 'base-cell-address', 'condition']);
        push(@{$style->{maps}}, $map);
      }
    }

  }

  return $style;
}

1;
