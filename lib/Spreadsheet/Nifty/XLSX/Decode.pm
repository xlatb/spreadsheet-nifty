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

# Given a string like "A1:C50", breaks it into fields.
sub decode2DRangeRefString($)
{
  my $class = shift();
  my ($str) = @_;

  ($str !~ m#^([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$#i) && return undef;

  my $ref = {};
  $ref->{minCol} = Spreadsheet::Nifty::Utils->stringToColIndex($1);
  $ref->{minRow} = int($2) - 1;
  $ref->{maxCol} = Spreadsheet::Nifty::Utils->stringToColIndex($3);
  $ref->{maxRow} = int($4) - 1;

  return $ref;
}

# As above, but we can have multiple ranges separated by spaces. The
#  result is always returned in an array.
sub decodeMulti2DRangeRefString($)
{
  my $class = shift();
  my ($str) = @_;

  my $parts = [ split(/ +/, $str) ];
  return [ map({ $class->decode2DRangeRefString($_) } @{$parts}) ];
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

  my $dims = $class->decode2DRangeRefString($ref);
  return $dims;
}

# <col min="1" max="1" width="18.7109375" customWidth="1"/>
sub decodeCol($)
{
  my $class = shift();
  my ($node) = @_;

  my $col = {};

  # Grab integer attributes
  for my $a (qw(min max customWidth bestFit style hidden outlineLevel collapsed))
  {
    my $v = $node->getAttribute($a);
    $col->{$a} = defined($v) ? int($v) : undef;
  }

  my $width = $node->getAttribute('width');
  $col->{width} = defined($width) ? (0 + $width) : undef;

  # The 'min' and 'max' attributes are 1-based, so convert to 0-based
  $col->{min}--;
  $col->{max}--;

  return $col;
}

# <row r="1000" ht="0" customHeight="1" hidden="1">
sub decodeRow($)
{
  my $class = shift();
  my ($node) = @_;

  my $row = {};

  $row->{rowIndex} = int($node->getAttribute('r') - 1);
  $row->{xfIndex} = $class->decodeBool($node->getAttribute('customFormat') // '0') ? $node->getAttribute('s') : undef;
  $row->{height} = $class->decodeBool($node->getAttribute('customHeight') // '0') ? $node->getAttribute('ht') : undef;
  $row->{hidden} = $class->decodeBool($node->getAttribute('hidden') // '0');

  return $row;
}

# <c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="A6" s="3" t="s"><v>5</v><f>4+1</f></c>
sub decodeCell($)
{
  my $class = shift();
  my ($node) = @_;

  my $cell = {};

  $cell->{type} = $node->getAttribute('t') // 'n';  # Optional attribute, defaults to numeric
  $cell->{xfIndex} = int($node->getAttribute('s') // 0);

  my $ref = $node->getAttribute('r');
  if (defined($ref))
  {
    ($ref !~ /^([A-Z]+)(\d+)$/) && die("Expected A1-style cell reference");
    $cell->{col} = Spreadsheet::Nifty::Utils->stringToColIndex($1);
    $cell->{row} = int($2) - 1;
  }

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
  if (defined($cell->{value}) && ($cell->{type} eq 'n'))
  {
    $cell->{value} = 0 + $cell->{value};  # NOTE: Should also work with scientific notation
  }

  # Include formula string, if any
  my ($f) = $node->getChildrenByTagNameNS($node->namespaceURI(), 'f');
  if (defined($f))
  {
    $cell->{formula} = $f->textContent();
  }

  return $cell;
}

# <autoFilter ref="A1:D100">...</autoFilter>
# Possible children are:
#  1. Zero or more <filterColumn/>
#  2. Optionally one <sortState/>
#  3. Optionally one <extList/>
sub decodeAutoFilter($)
{
  my $class = shift();
  my ($node) = @_;

  my $autoFilter = {};
  $autoFilter->{ref} = $class->decode2DRangeRefString($node->getAttribute('ref'));

  $autoFilter->{filterColumns} = [];
  my $filterColumnElements = [ $node->getChildrenByTagNameNS($node->namespaceURI(), 'filterColumn') ];
  for my $e (@{$filterColumnElements})
  {
    push(@{$autoFilter->{filterColumns}}, $class->decodeFilterColumn($e));
  }

  my ($sortStateElement) = $node->getChildrenByTagNameNS($node->namespaceURI(), 'sortState');
  if (defined($sortStateElement))
  {
    $autoFilter->{sortState} = $class->decodeSortState($sortStateElement);
  }

  return $autoFilter;
}

# <filterColumn colId="0">...</filterColumn>
# If there are active filters, we will have one of the following children:
# 1. <filters/>
# 2. <top10/>
# 3. <customFilters/>
# 4. <dynamicFilters/>
# 5. <colorFilter/>
# 6. <iconFilter/>
sub decodeFilterColumn($)
{
  my $class = shift();
  my ($node) = @_;

  my $filterColumn = {};
  $filterColumn->{colId} = int($node->getAttribute('colId'));
  $filterColumn->{hiddenButton} = int($node->getAttribute('hiddenButton') // 0);
  $filterColumn->{showButton} = int($node->getAttribute('showButton') // 1);

  if (defined(my ($filtersElement) = $node->getChildrenByTagNameNS($node->namespaceURI(), 'filters')))
  {
    $filterColumn->{filterType} = 'filters';
    $filterColumn->{filters} = $class->decodeFilters($filtersElement);
  }
  elsif (defined(my ($customFiltersElement) = $node->getChildrenByTagNameNS($node->namespaceURI(), 'customFilters')))
  {
    $filterColumn->{filterType} = 'customFilters';
    $filterColumn->{filters} = $class->decodeCustomFilters($customFiltersElement);
  }
  # TODO: Other filter types

  return $filterColumn;
}

# <filters blank="1">
#   <filter val="ABC"/>
#   <filter val="123"/>
#   <dateGroupItem year="2026" month="11" dateTimeGrouping="month"/>
# </filters>
sub decodeFilters($)
{
  my $class = shift();
  my ($node) = @_;

  my $filters = {};
  $filters->{blank} = int($node->getAttribute('blank') // 0);  # Whether blank cells are shown
  $filters->{calendarType} = $node->getAttribute('calendarType');

  my $filterElements = [ $node->getChildrenByTagNameNS($node->namespaceURI(), 'filter') ];
  $filters->{values} = [ map({ $_->getAttribute('val') } @{$filterElements}) ];

  # TODO: dateGroupItem

  return $filters;
}

# <customFilters and="1">
#   <customFilter operator="greaterThanOrEqual" val="10"/>
#   <customFilter operator="lessThanOrEqual" val="50"/>
# <customFilters>
# The maximum number of customFilter children is 2, to allow for a range.
sub decodeCustomFilters($)
{
  my $class = shift();
  my ($node) = @_;

  my $customFilters = {};
  $customFilters->{and} = int($node->getAttribute('blank') // 0);  # AND vs OR
  $customFilters->{filters} = [];

  my $customFilterElements = [ $node->getChildrenByTagNameNS($node->namespaceURI(), 'customFilter') ];
  for my $e (@{$customFilterElements})
  {
    my $customFilter = {};
    $customFilter->{operator} = $e->getAttribute('operator') // 'equal';
    $customFilter->{val} = $e->getAttribute('val');
    push(@{$customFilters->{filters}}, $customFilter);
  }

  return $customFilters;
}

# <sortState ref="A2:C50">...</sortState>
sub decodeSortState($)
{
  my $class = shift();
  my ($node) = @_;

  my $sortState = {};
  $sortState->{ref} = $class->decode2DRangeRefString($node->getAttribute('ref'));
  $sortState->{columnSort} = int($node->getAttribute('columnSort') // 0);
  $sortState->{caseSensitive} = int($node->getAttribute('caseSensitive') // 0);
  $sortState->{sortMethod} = $node->getAttribute('sortMethod');

  my $sortConditionElements = [ $node->getChildrenByTagNameNS($node->namespaceURI(), 'sortCondition') ];
  $sortState->{conditions} = [];
  for my $e (@{$sortConditionElements})
  {
    push(@{$sortState->{conditions}}, $class->decodeSortcondition($e));
  }

  return $sortState;
}

# <sortCondition ref="B2:B50" descending="1"/>
sub decodeSortCondition($)
{
  my $class = shift();
  my ($node) = @_;

  my $sortCondition = {};
  $sortCondition->{ref} = $class->decode2DRangeRefString($node->getAttribute('ref'));
  $sortCondition->{descending} = int($node->getAttribute('descending') // 0);
  $sortCondition->{sortBy} = $node->getAttribute('sortBy') // 'value';
  $sortCondition->{customList} = $node->hasAttribute('customList') ? [ split(/,/, $node->getAttribute('customList')) ] : undef;
  $sortCondition->{dxfId} = $node->hasAttribute('dxfId') ? int($node->getAttribute('dxfId')) : undef;
  $sortCondition->{iconSet} = $node->getAttribute('iconSet');
  $sortCondition->{iconId} = node->hasAttribute('iconId') ? int($node->getAttribute('iconId')) : undef;
  return $sortCondition;
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
