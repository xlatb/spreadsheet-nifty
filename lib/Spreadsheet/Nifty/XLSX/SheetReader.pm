#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSX::SheetReader;

use Spreadsheet::Nifty::XLSX;
use Spreadsheet::Nifty::XLSX::Decode;
use Spreadsheet::Nifty::ZIPPackage;
use Spreadsheet::Nifty::IndexedColors;

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
  $self->{header}         = undef;
  $self->{sharedFormulae} = {};

  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub open()
{
  my $self = shift();

  $self->{xmlReader} = XML::LibXML::Reader->new({IO => $self->{zipReader}});

  $self->{header}     = {};
  $self->{rows}       = [];
  $self->{rowIndices} = {next => 0, final => undef, highest => undef};

  return $self->readHeader();
}

sub rewind()
{
  my $self = shift();

  $self->{xmlReader}->close();
  $self->{zipReader}->seek(0);

  return $self->open();
}

# Reads up until <sheetData/>.
sub readHeader()
{
  my $self = shift();

  my $xmlns = $Spreadsheet::Nifty::XLSX::namespaces->{main};
  
  # Read elements up to <sheetData/>. Apparently they are all optional.
  my $status;
  while (($status = $self->{xmlReader}->nextElement()) == 1)
  {
    #$self->{debug} && printf("  element %s:%s\n", $self->{xmlReader}->namespaceURI(), $self->{xmlReader}->localName());
    ($self->{xmlReader}->namespaceURI() ne $xmlns) && next;
    ($self->{xmlReader}->nodeType == XML_READER_TYPE_END_ELEMENT) && next;  # Don't process element end tags

    my $localName = $self->{xmlReader}->localName();
    if ($localName eq 'sheetData')
    {
      # TODO: Skip over this element and continue reading? Stuff like autoFilter and sheetProtection occurs after
      $self->{xmlReader}->read();  # Consume <sheetData/>
      $self->{rowIndices}->{next} = 0;
      return 1;
    }
    elsif ($localName eq 'dimension')
    {
      my $node = $self->{xmlReader}->copyCurrentNode(0);
      $self->{header}->{dimensions} = Spreadsheet::Nifty::XLSX::Decode->decodeDimension($node);
    }
    elsif ($localName eq 'sheetProtection')
    {
      my $node = $self->{xmlReader}->copyCurrentNode(1);
      $self->{header}->{protection} = Spreadsheet::Nifty::XLSX::Decode->decodeWorksheetProtection($node);
    }
  }

  # Hit end of document without seeing <sheetData/>.
  return 0;
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

  my $xmlns = $Spreadsheet::Nifty::XLSX::namespaces->{main};

  ($self->{debug}) && printf("findRow(): rowIndex %d\n", $rowIndex);

  # Row index cannot be negative
  ($rowIndex < 0) && return undef;

  # If header included dimensions, check upper bound
  if (defined($self->{header}->{dimensions}))
  {
    ($rowIndex > $self->{header}->{dimensions}->{maxRow}) && return 0;
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
  
  my $xmlns = $Spreadsheet::Nifty::XLSX::namespaces->{main};

  if (($self->{xmlReader}->namespaceURI() ne $xmlns) || ($self->{xmlReader}->localName() ne 'row'))
  {
    die("readRow(): Expected to start reading at <row/> element.");
  }

  my $row = Spreadsheet::Nifty::XLSX::Decode->decodeRow($self->{xmlReader}->copyCurrentNode(0));
  $self->{xmlReader}->nextElement();
  
  my $cells = [];
  while (1)
  {
    #printf("readRow() depth %d nodeType %d localname %s\n", $self->{xmlReader}->depth(), $self->{xmlReader}->nodeType(), $self->{xmlReader}->localName());

    if ($self->{xmlReader}->depth() < 3)
    {
      # No longer within <row/>
      last;
    }

    if (($self->{xmlReader}->nodeType() == XML_READER_TYPE_ELEMENT) && ($self->{xmlReader}->namespaceURI() eq $xmlns) && ($self->{xmlReader}->localName() eq 'c'))
    {
      my $node = $self->{xmlReader}->copyCurrentNode(1);
      my $cell = Spreadsheet::Nifty::XLSX::Decode->decodeCell($node);

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

    my $status = $self->{xmlReader}->nextSibling();
    ($status != 1) && last;  # No more elements
  }

  # Elements with start and end tags come up twice, once for the start tag and
  #  once for the end tag. If we're positioned on a row end tag, consume it.
  if (($self->{xmlReader}->nodeType() == XML_READER_TYPE_END_ELEMENT) && ($self->{xmlReader}->namespaceURI() eq $xmlns) && ($self->{xmlReader}->localName() eq 'row'))
  {
    $self->{xmlReader}->nextSibling();
  }

  return $cells;
}

1;
