#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSX::Decode;

use Spreadsheet::Nifty::Utils;

use XML::LibXML qw(:libxml);

# Given a DOM element, returns the attributes as a hash reference.
# If the DOM element is undefined, returns undef.
sub attributesHash($)
{
  my $class = shift();
  my ($element) = @_;

  (!defined($element)) && return undef;

  my $attributes = [ $element->attributes() ];
  my $hash = {};

  for my $a (@{$attributes})
  {
    $hash->{$a->nodeName()} = $a->value();
  }

  return $hash;
}

# Given an XML Schema boolean value string, returns the boolean value.
sub decodeBool($)
{
  my $class = shift();
  my ($str) = @_;

  (!defined($str)) && return undef;

  if (($str eq '1') || ($str eq 'true'))
  {
    return !!1;
  }
  elsif (($str eq '0') || ($str eq 'false'))
  {
    return !!0;
  }

  return undef;
}

# Given a DOM node which is a string container, extracts the string contained
#  within it. Valid nodes to pass are <is/>, <r/>, <rPh/>, <si/>, or <text/>.
sub decodePlainString($)
{
  my $class = shift();
  my ($node) = @_;
  
  my $xmlns = $Spreadsheet::Nifty::XLSX::namespaces->{main};
  
  my $string = '';
  
  for (my $kid = $node->firstChild(); defined($kid); $kid = $kid->nextSibling())
  {
    ($kid->nodeType() == XML_ELEMENT_NODE) || next;
    ($kid->namespaceURI() eq $xmlns) || next;
    
    my $name = $kid->localname();
    if ($name eq 't')
    {
      # Text run
      $string .= $kid->textContent();
    }
    elsif ($name eq 'r')
    {
      # Rich text run, recurse
      $string .= $class->decodePlainString($kid);
    }
  }
  
  return $string;
}

# <dimension ref="A2:T22" />
sub decodeDimension($)
{
  my $class = shift();
  my ($node) = @_;

  my $ref = $node->getAttribute('ref');
  (!defined($ref)) && return undef;

  ($ref !~ m#^([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$#i) && return undef;

  my $dims = {};
  $dims->{minCol} = Spreadsheet::Nifty::Utils->stringToColIndex($1);
  $dims->{minRow} = int($2);
  $dims->{maxCol} = Spreadsheet::Nifty::Utils->stringToColIndex($3);
  $dims->{maxRow} = int($4);

  return $dims;
}

# <row r="1000" ht="0" customHeight="1" hidden="1">
sub decodeRow($)
{
  my $class = shift();
  my ($node) = @_;

  my $row = {};

  $row->{rowIndex} = int($node->getAttribute('r') - 1);
  $row->{styleIndex} = $class->decodeBool($node->getAttribute('customFormat') // '0') ? $node->getAttribute('s') : undef;
  $row->{height} = $class->decodeBool($node->getAttribute('customHeight') // '0') ? $node->getAttribute('ht') : undef;
  $row->{hidden} = $class->decodeBool($node->getAttribute('hidden') // '0');

  return $row;
}

# <c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="A6" s="3" t="s"><v>5</v></c>
sub decodeCell($)
{
  my $class = shift();
  my ($node) = @_;

  my $cell = {};

  $cell->{type} = $node->getAttribute('t') // 'n';  # Optional attribute, defaults to numeric
  $cell->{styleIndex} = int($node->getAttribute('s') // 0);

  my $ref = $node->getAttribute('r');
  if (defined($ref))
  {
    ($ref !~ /^([A-Z]+)(\d+)$/) && die("Expected A1-style cell reference");
    $cell->{col} = Spreadsheet::Nifty::Utils->stringToColIndex($1);
    $cell->{row} = int($2) - 1;
  }

  my $value = undef;
  my ($is) = $node->getChildrenByTagNameNS($node->namespaceURI(), 'is');
  if (defined($is) && (!defined($cell->{type}) || ($cell->{type} eq 'inlineStr')))
  {
    # Element <is/> as direct descendant of <c/> (section 18.3.1.53). Cell
    #  type of 'inlineStr' is apparently optional.
    $cell->{value} = Spreadsheet::Nifty::XLSX::Decode->decodePlainString($is);
  }
  else
  {
    my ($v) = $node->getChildrenByTagNameNS($node->namespaceURI(), 'v');
    if (defined($v))
    {
      if ($cell->{type} eq 's')
      {
        $cell->{stringIndex} = int($v->textContent());
      }
      else
      {
        $cell->{value} = $v->textContent();
      }
    }
  }
      
  # Coerce numeric values to numbers
  if (defined($value) && ($cell->{type} eq 'n'))
  {
    $cell->{value} = 0 + $cell->{value};  # NOTE: Should also work with scientific notation
  }

  return $cell;
}

# <workbookProtection/> element for an entire workbook.
# If 'type' is defined, the workbook is protected.
sub decodeWorkbookProtection($)
{
  my $class = shift();
  my ($node) = @_;

  my $protection = {};

  my $algo = $node->getAttribute('workbookAlgorithmName');
  if (defined($algo))
  {
    # New style
    $protection->{type}      = $algo;
    $protection->{hash}      = $node->getAttribute('workbookHashValue');
    $protection->{salt}      = $node->getAttribute('workbookSaltValue');
    $protection->{spinCount} = $node->getAttribute('workbookSpinCount');
  }
  elsif (defined(my $password = $node->getAttribute('workbookPassword')))
  {
    # Old style
    $protection->{type} = 'excel16';
    $protection->{hash} = hex($password);
  }

  # Parse restrictions
  $protection->{restrictions} = {};
  for my $n (qw(lockRevision lockStructure lockWindows))
  {
    my $v = $node->getAttribute($n);
    $protection->{restrictions}->{$n} = $v;
  }

  return $protection;
}

# <sheetProtection/> element for a single worksheet.
# If 'type' is defined, the workbook is protected.
sub decodeWorksheetProtection($)
{
  my $class = shift();
  my ($node) = @_;

  my $protection = {};

  my $algo = $node->getAttribute('algorithmName');
  if (defined($algo))
  {
    # New style
    $protection->{type}      = $algo;
    $protection->{hash}      = $node->getAttribute('hashValue');
    $protection->{salt}      = $node->getAttribute('saltValue');
    $protection->{spinCount} = $node->getAttribute('spinCount');
  }
  elsif (defined(my $password = $node->getAttribute('password')))
  {
    # Old style
    $protection->{type} = 'excel16';
    $protection->{hash} = hex($password);
  }

  my $names =
  [
    'sheet',
    'autoFilter', 'sort',
    'pivotTables',
    'insertRows', 'deleteRows', 'formatRows',
    'insertColumns', 'deleteColumns', 'formatColumns',
    'insertHyperlinks',
    'selectLockedCells', 'selectUnlockedCells',
  ];

  # Parse restrictions
  $protection->{restrictions} = {};
  for my $n (@{$names})
  {
    my $v = $node->getAttribute($n);
    $protection->{restrictions}->{$n} = $v;
  }

  return $protection;
}

1;
