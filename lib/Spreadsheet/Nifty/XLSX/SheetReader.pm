#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSX::SheetReader;

my $xmlns = $Spreadsheet::Nifty::XLSX::namespaces->{main};

use Spreadsheet::Nifty::XLSX;
use Spreadsheet::Nifty::XLSX::Decode;
use Spreadsheet::Nifty::ZIPPackage;
use Spreadsheet::Nifty::IndexedColors;
use Spreadsheet::Nifty::XMLReaderUtils;

use XML::LibXML qw(:libxml);
use XML::LibXML::Reader;
use Data::Dumper;

# === Class methods ===

sub new()
{
  my $class = shift();
  my ($workbook, $sheetIndex, $zipReader) = @_;
  
  my $self = {};
  $self->{workbook}       = $workbook;
  $self->{sheetIndex}     = $sheetIndex;
  $self->{debug}          = 0;
  $self->{zipReader}      = $zipReader;
  $self->{xmlReader}      = undef;
  $self->{rowIndices}     = {};
  $self->{rows}           = undef;
  $self->{meta}           = {};
  $self->{sharedFormulae} = {};

  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub workbook()
{
  my $self = shift();

  return $self->{workbook};
}

sub init()
{
  my $self = shift();

  $self->{xmlReader} = XML::LibXML::Reader->new({IO => $self->{zipReader}});

  $self->{rows}       = [];
  $self->{rowIndices} = {next => 0, final => undef, highest => undef};

  return !!1;
}

sub open()
{
  my $self = shift();

  (!$self->init()) && return !!0;
  (!$self->readRootElement()) && return !!0;
  (!$self->readPrologue()) && return !!0;

  return !!1;
}

# Rewind the stream to the beginning.
# Note that most methods assume that the read position is left within
#  the <sheetData/> element. It's better to use findSheetDataStart()
#  if seeking backwards is the goal.
sub rewind()
{
  my $self = shift();

  $self->{xmlReader}->close();
  $self->{zipReader}->seek(0);

  return $self->init();
}

sub findSheetDataStart()
{
  my $self = shift();

  (!$self->rewind()) && return !!0;
  (!$self->readRootElement()) && return !!0;

  my $status;
  do
  {
    #$self->{debug} && printf("  element %s:%s\n", $self->{xmlReader}->namespaceURI(), $self->{xmlReader}->localName());
    ($self->{xmlReader}->namespaceURI() ne $xmlns) && next;
    ($self->{xmlReader}->nodeType == XML_READER_TYPE_END_ELEMENT) && next;  # Don't process element end tags

    my $localName = $self->{xmlReader}->localName();
    if ($localName eq 'sheetData')
    {
      $self->{xmlReader}->read();  # Descend into <sheetData/>
      $self->{rowIndices}->{next} = 0;
      return !!1;
    }
  } while (($status = $self->{xmlReader}->nextSiblingElement()) == 1);

  return !!0;  # Never found <sheetData/>
}

sub readRootElement()
{
  my $self = shift();

  ($self->{xmlReader}->readState() != XML_READER_START) && die("readRootElement(): Not at start of stream");

  # Check the root element
  $self->{xmlReader}->nextElement();
  ($self->{xmlReader}->localName() ne 'worksheet') && return !!0;
  ($self->{xmlReader}->namespaceURI() ne $xmlns) && return !!0;

  # Descend into its children
  $self->{xmlReader}->nextElement();

  return !!1;
}

# Reads the prologue information appearing before <sheetData/>.
# The read location is left at the start of the contents of <sheetData/>.
sub readPrologue()
{
  my $self = shift();

  # Read child elements up to <sheetData/>. Apparently they are all optional.
  my $status;
  do
  {
    #$self->{debug} && printf("  element %s:%s\n", $self->{xmlReader}->namespaceURI(), $self->{xmlReader}->localName());
    ($self->{xmlReader}->namespaceURI() ne $xmlns) && next;
    ($self->{xmlReader}->nodeType == XML_READER_TYPE_END_ELEMENT) && next;  # Don't process element end tags

    my $localName = $self->{xmlReader}->localName();
    if ($localName eq 'sheetData')
    {
      $self->{xmlReader}->read();  # Descend into <sheetData/>
      $self->{rowIndices}->{next} = 0;
      return 1;
    }
    elsif ($localName eq 'dimension')
    {
      my $node = $self->{xmlReader}->copyCurrentNode(0);
      $self->{meta}->{dimensions} = Spreadsheet::Nifty::XLSX::Decode->decodeDimension($node);
    }
    elsif ($localName eq 'cols')
    {
      $self->{meta}->{columns} = $self->readColumnAttributes();
    }
  } while (($status = $self->{xmlReader}->nextSiblingElement()) == 1);

  # Hit end of document without seeing <sheetData/>.
  return !!0;
}

sub readEpilogueIfNeeded()
{
  my $self = shift();

  ($self->{meta}->{epilogueRead}) && return !!1;  # Already read it

  return $self->readEpilogue();
}

# Reads the epilogue information appearing after <sheetData/>.
# We rewind the stream afterwards so that we are ready to read sheet data.
sub readEpilogue()
{
  my $self = shift();

  # Find the end tag of <sheetData/> if we are not already there
  if (!Spreadsheet::Nifty::XMLReaderUtils->atEndOfElement($self->{xmlReader}, 'sheetData', $xmlns))
  {
    (!Spreadsheet::Nifty::XMLReaderUtils->ascendToElement($self->{xmlReader}, 'sheetData', $xmlns)) && die("readEpilogue(): Could not find end of <sheetData/>");
  }

  # Read the elements following <sheetData/>.
  my $status;
  while (($status = $self->{xmlReader}->nextSiblingElement()) == 1)
  {
    ($self->{xmlReader}->namespaceURI() ne $xmlns) && next;
    ($self->{xmlReader}->nodeType == XML_READER_TYPE_END_ELEMENT) && next;  # Don't process element end tags

    my $localName = $self->{xmlReader}->localName();
    if ($localName eq 'autoFilter')
    {
      my $node = $self->{xmlReader}->copyCurrentNode(1);
      my $autoFilter = Spreadsheet::Nifty::XLSX::Decode->decodeAutoFilter($node);
      $self->{meta}->{autoFilter} = $autoFilter;
    }
    elsif ($localName eq 'dataValidations')
    {
      my $node = $self->{xmlReader}->copyCurrentNode(1);
      $self->{meta}->{dataValidations} = [];
      for my $dv ($node->getChildrenByTagNameNS($node->namespaceURI(), 'dataValidation'))
      {
        push(@{$self->{meta}->{dataValidations}}, Spreadsheet::Nifty::XLSX::Decode->decodeDataValidation($dv));
      }
    }
    elsif ($localName eq 'sheetProtection')
    {
      my $node = $self->{xmlReader}->copyCurrentNode(1);
      $self->{meta}->{protection} = Spreadsheet::Nifty::XLSX::Decode->decodeWorksheetProtection($node);
    }
    elsif ($localName eq 'extLst')
    {
      # Process each extension within extLst
      if (Spreadsheet::Nifty::XMLReaderUtils->findChildElement($self->{xmlReader}, 'ext', $xmlns))
      {
        do
        {
          my $ext = $self->{xmlReader}->copyCurrentNode(1);
          my $uri = uc($ext->getAttribute('uri'));
          if ($uri eq $Spreadsheet::Nifty::XLSX::extensionIds->{'x14:dataValidations'})
          {
            my ($dataValidationsX14) = $ext->getChildrenByTagNameNS($Spreadsheet::Nifty::XLSX::namespaces->{x14}, 'dataValidations');
            for my $dvX14 ($dataValidationsX14->getChildrenByTagNameNS($Spreadsheet::Nifty::XLSX::namespaces->{x14}, 'dataValidation'))
            {
              push(@{$self->{meta}->{dataValidations}}, Spreadsheet::Nifty::XLSX::Decode->decodeDataValidationX14($dvX14));
            }
          }
        } while (Spreadsheet::Nifty::XMLReaderUtils->findSiblingElement($self->{xmlReader}, 'ext', $xmlns));
      }
    }
  }

  (!$self->rewind()) && die("readEpilogue(): rewind failed");
  (!$self->findSheetDataStart()) && die("readEpilogue(): findSheetDataStart failed");

  $self->{meta}->{epilogueRead} = !!1;
  return !!1;
}

sub readColumnAttributes()
{
  my $self = shift();

  if (($self->{xmlReader}->namespaceURI() ne $xmlns) || ($self->{xmlReader}->localName() ne 'cols'))
  {
    die("readColumnAttributes(): Expected to start reading at <cols/> element.");
  }

  my $colElements = [];
  if (Spreadsheet::Nifty::XMLReaderUtils->findChildElement($self->{xmlReader}, 'col', $xmlns))
  {
    do
    {
      my $node = $self->{xmlReader}->copyCurrentNode(1);
      push(@{$colElements}, Spreadsheet::Nifty::XLSX::Decode->decodeCol($node));
    } while (Spreadsheet::Nifty::XMLReaderUtils->findSiblingElement($self->{xmlReader}, 'col', $xmlns));
  }

  # Repeated column definitions only get a single <col/> element, so we expand
  #  them here. Note that all instances of the repeated column reference the
  #  same shared data.
  my $colAttrs = [];
  for my $ce (@{$colElements})
  {
    for (my $i = $ce->{min}; $i <= $ce->{max}; $i++)
    {
      $colAttrs->[$i] = $ce;
    }
  }

  return $colAttrs;
}

sub parseFormula($$$$)
{
  my $self = shift();
  my ($f, $row, $col, $sharedFormulae) = @_;

  # If 't' attribute is 'shared', we will have a share index (si). First
  #  instance is the "master formula" and has 'ref' attribute:
  #   <f t="shared" ref="H7:H11" ce="1" si="0">SUM(E7:G7)</f>
  #
  # Subsequent formulae may reference this by having 'si' but no 'ref':
  #   <f t="shared" ce="1" si="0">SUM(E8:G8)</f>
  #
  # Note that, as in this example, the formula *may* re-appear in references,
  #  but it is optional.

  my $t = $f->getAttribute('t') // 'normal';
  if (($t eq 'normal') || ($t eq 'array') || ($t eq 'dataTable'))
  {
    my $formula = {t => $t, formula => $f->textContent()};
    return $formula;
  }
  elsif ($t eq 'shared')
  {
    my $si  = int($f->getAttribute('si'));
    my $ref = $f->getAttribute('ref');
    if (defined($ref))
    {
      my $formula = {t => $t, defineRow => $row, defineCol => $col, formula => $f->textContent};
      $sharedFormulae->{$si} = $formula;
      return $formula;
    }
    elsif (defined($sharedFormulae->{$si}))
    {
      # Return shared formula. Note that the string in the 'formula' field
      #  needs to be rewritten to be accurate for a new cell. We don't bother
      #  to do that here for efficiency reasons.
      return $sharedFormulae->{$si};
    }
    die("Reference to unknown shared formula");
  }

  die("Unknown formula type");
}

# Given a row index, positions our reader at the <row/> element if possible.
# Returns undef if no such row exists but it's still possible that higher-numbered rows might.
# Returns 0 if the requested row is beyond the end of the file.
# Assumes our read position is within <sheetData/>.
sub findRow($)
{
  my $self = shift();
  my ($rowIndex) = @_;

  ($self->{debug}) && printf("findRow(): rowIndex %d\n", $rowIndex);

  # Row index cannot be negative
  ($rowIndex < 0) && return undef;

  # If prologue included dimensions, check upper bound
  if (defined($self->{meta}->{dimensions}))
  {
    ($rowIndex > $self->{meta}->{dimensions}->{maxRow}) && return 0;
  }

  # If final row is known, check upper bound
  if (defined($self->{rowIndices}->{final}))
  {
    ($rowIndex > $self->{rowIndices}->{final}) && return 0;
  }

  # If requested row index is behind us, we need to rewind
  if ($rowIndex < $self->{rowIndices}->{next})
  {
    #printf("  Rewinding...\n");
    (!$self->rewind()) && return undef;  # Couldn't rewind
  }

  # Read elements within <sheetData/>. Each should be a <row/>.
  while (1)
  {
    #printf("findRow() depth %d nodeType %d localname %s\n", $self->{xmlReader}->depth(), $self->{xmlReader}->nodeType(), $self->{xmlReader}->localName());

    if ($self->{xmlReader}->depth() < 2)
    {
      # No longer within <sheetData/>
      last;
    }

    if (($self->{xmlReader}->nodeType() == XML_READER_TYPE_ELEMENT) && ($self->{xmlReader}->namespaceURI() eq $xmlns) && ($self->{xmlReader}->localName() eq 'row'))
    {
      my $node = $self->{xmlReader}->copyCurrentNode(0);
      my $row = Spreadsheet::Nifty::XLSX::Decode->decodeRow($node);
      $self->{rows}->[$row->{rowIndex}] = $row;

      # Track highest row number seen so far
      if (!defined($self->{rowIndices}->{highest}) || ($row->{rowIndex} > $self->{rowIndices}->{highest}))
      { 
        $self->{rowIndices}->{highest} = $row->{rowIndex};
      }

      $self->{rowIndices}->{next} = $row->{rowIndex} + 1;
      #printf("  Saw row %d nodeType %d: %s\n", $row->{rowIndex}, $self->{xmlReader}->nodeType(), $node->toString());

      if ($rowIndex == $row->{rowIndex})
      {
        #printf("    Found target row %d\n", $rowIndex);
        return $row;
      }
      elsif ($rowIndex < $row->{rowIndex})
      {
        #printf("    Reached row %d which is beyond target row %d\n", $row->{rowIndex}, $rowIndex);
        return undef;
      }
    }

    #my $status = $self->{xmlReader}->nextElement('row', $xmlns);
    my $status = $self->{xmlReader}->nextSibling();
    ($status != 1) && last;  # No more row elements
  };

  # We fell off the end of the row data, so record final number
  #printf("    Reached end of rows searching for target row %d. Highest seen was %d.\n", $rowIndex, $self->{rowIndices}->{highest} // -1);
  $self->{rowIndices}->{final} = $self->{rowIndices}->{highest} // -1;
  return 0;
}

sub readRow($)
{
  my $self = shift();
  
  if (($self->{xmlReader}->namespaceURI() ne $xmlns) || ($self->{xmlReader}->localName() ne 'row'))
  {
    die("readRow(): Expected to start reading at <row/> element.");
  }

  my $rowNode = $self->{xmlReader}->copyCurrentNode(1);
  my $row = Spreadsheet::Nifty::XLSX::Decode->decodeRow($rowNode);
  $self->{xmlReader}->nextSibling();

  # Some XLSX spreadsheets suffer from a "row bloat" problem, where an entire
  #  row had a format applied, and in response Excel manifests every cell
  #  in the row as an empty cell with a style applied to it.
  # This makes extracting the row data very slow because we're needlessly 
  #  churniing through 16k cells per row.
  # We use this xpath to find the final cell in the row that contains data.
  my $xpath = XML::LibXML::XPathContext->new($rowNode);
  $xpath->registerNs('excel', $xmlns);
  my ($final) = $xpath->findnodes('child::excel:c[excel:v or excel:is][last()]');
  ($self->{debug}) && printf("Final cell with a value on this row: %s\n", $final ? $final->toString() : '(undef)');
  (!defined($final)) && return [];  # Quick exit if there are no values on this row

  # Now we retrieve the preceding cells on the row, whether or not they
  #  contain a value, and use that to construct our array of cells to process
  #  on this row.
  my $cNodes = [ $xpath->findnodes('preceding-sibling::excel:c', $final) ];
  push(@{$cNodes}, $final);  # Add back the final one

  my $cells = [];
  for my $cNode (@{$cNodes})
  {
    my $cell = Spreadsheet::Nifty::XLSX::Decode->decodeCell($cNode);

    (defined($cell->{row}) && ($row->{rowIndex} != $cell->{row})) && die("Read a cell from the wrong row?");

    if (defined($cell->{col}))
    {
      $cells->[$cell->{col}] = $cell;
    }
    else
    {
      push(@{$cells}, $cell);
    }
  }

  # Elements with start and end tags come up twice, once for the start tag and
  #  once for the end tag. If we're positioned on a row end tag, consume it.
  if (($self->{xmlReader}->nodeType() == XML_READER_TYPE_END_ELEMENT) && ($self->{xmlReader}->namespaceURI() eq $xmlns) && ($self->{xmlReader}->localName() eq 'row'))
  {
    $self->{xmlReader}->nextSibling();
  }

  return $cells;
}

sub getAutoFilterState()
{
  my $self = shift();

  $self->readEpilogueIfNeeded();

  (!defined($self->{meta}->{autoFilter})) && return undef;

  my $autoFilter = {};
  $autoFilter->{range} = { %{$self->{meta}->{autoFilter}->{ref}} };

  return $autoFilter;
}

# Returns all data validations
sub getDataValidations()
{
  my $self = shift();

  $self->readEpilogueIfNeeded();

  return $self->{meta}->{dataValidations};
}

# Returns data validations that cover the given cell
sub getDataValidationsForCell($$)
{
  my $self = shift();
  my ($r, $c) = @_;

  $self->readEpilogueIfNeeded();

  (!defined($self->{meta}->{dataValidations})) && return undef;

  my $dataValidations = [];
  for my $dv (@{$self->{meta}->{dataValidations}})
  {
    my $matched = !!0;
    for my $ref (@{$dv->{refs}})
    {
      (($c < $ref->{minCol}) || ($c > $ref->{maxCol})) && next;
      (($r < $ref->{minRow}) || ($r > $ref->{maxRow})) && next;
      $matched = !!1;
      last;
    }
    (!$matched) && next;
    push(@{$dataValidations}, $dv);
  }

  return $dataValidations;
}

1;

