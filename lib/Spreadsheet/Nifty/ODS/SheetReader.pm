#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::ODS::SheetReader;

use Spreadsheet::Nifty::OpenDocument;
use Spreadsheet::Nifty::XMLReaderUtils;
use Spreadsheet::Nifty::ODS;
use Spreadsheet::Nifty::ODS::Decode;

use XML::LibXML qw(:libxml);
use XML::LibXML::Reader;
use Data::Dumper;

# === Class methods ===

sub new($$)
{
  my $class = shift();
  my ($workbook, $sheetIndex) = @_;
  
  my $self = {};
  $self->{workbook}   = $workbook;
  $self->{sheetIndex} = $sheetIndex;
  $self->{xmlReader}  = undef;
  $self->{rowIndex}   = 0;
  $self->{rowCount}   = $workbook->getSheetRowCount($sheetIndex);
  $self->{currentRow} = undef;
  $self->{columnDefs} = undef;
  $self->{rowDefs}    = undef;
  $self->{debug}      = 0;
  
  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub open()
{
  my $self = shift();

  my $xmlReader = $self->{workbook}->xmlReaderForSheet($self->{sheetIndex});
  (!defined($xmlReader)) && return !!0;

  $self->{xmlReader} = $xmlReader;

  return $self->readInitial();
}

# Read until the first <table:table-row/> is seen.
sub readInitial()
{
  my $self = shift();

  my $xmlns = $Spreadsheet::Nifty::ODS::namespaces->{table};
  
  # Sanity check reader position
  (!Spreadsheet::Nifty::XMLReaderUtils->atStartOfElement($self->{xmlReader}, 'table', $xmlns)) && return !!0;

  # Step into children
  ($self->{xmlReader}->isEmptyElement()) && return !!0;
  my $status = $self->{xmlReader}->read();

  while ($status == 1)
  {
    #Spreadsheet::Nifty::XMLReaderUtils->dump($self->{xmlReader}, 'readInitial()');

    my $localName = $self->{xmlReader}->localName();
    my $ns = $self->{xmlReader}->namespaceURI();

    if ($ns eq $xmlns)
    {
      if ($localName eq 'table-column')
      {
        $self->readColumns();
        next;
      }
      elsif ($localName eq 'table-row')
      {
        $self->{rowDefs} = [];
        return !!1;  # Reached first <:table:table-row/>
      }
    }

    # Something unknown, skip it
    #printf("  Skipping '%s'...\n", $self->{xmlReader}->localName());
    Spreadsheet::Nifty::XMLReaderUtils->skipElement($self->{xmlReader});
    $status = $self->{xmlReader}->read();
  }
}

sub readColumns()
{
  my $self = shift();

  #printf("readColumns()\n");

  my $xmlns = $Spreadsheet::Nifty::ODS::namespaces->{table};
  
  my $columnDefs = [];
  while (Spreadsheet::Nifty::XMLReaderUtils->atStartOfElement($self->{xmlReader}, 'table-column', $xmlns))
  {
    #Spreadsheet::Nifty::XMLReaderUtils->dump($self->{xmlReader}, "readColumns()");
    my $columnDef = Spreadsheet::Nifty::ODS::Decode->decodeColumnDefinition($self->{xmlReader}->copyCurrentNode(1));
    push(@{$columnDefs}, $columnDef);
    ($self->{xmlReader}->read() != 1) && die("readColumns(): XML read error");
  }

  $self->{columnDefs} = $columnDefs;
  #print main::Dumper($columnDefs);

  return;
}

sub countColumns()
{
  my $self = shift();

  my $count = 0;
  for my $cd (@{$self->{columnDefs}})
  {
    $count += $cd->{count};
  }

  return $count;
}

sub parkXmlReader()
{
  my $self = shift();

  my $xmlns = $Spreadsheet::Nifty::ODS::namespaces->{table};

  # Move reader to the end of the table
  if (!Spreadsheet::Nifty::XMLReaderUtils->atEndOfElement($self->{xmlReader}, 'table', $xmlns))
  {
    Spreadsheet::Nifty::XMLReaderUtils->ascendToElement($self->{xmlReader}, 'table', $xmlns);
  }

  # Park it
  $self->{workbook}->parkXmlReaderForSheet($self->{xmlReader}, $self->{sheetIndex});
  $self->{xmlReader} = undef;

  return;
}

sub readRow()
{
  my $self = shift();

  my $xmlns = $Spreadsheet::Nifty::ODS::namespaces->{table};
  #Spreadsheet::Nifty::XMLReaderUtils->dump($self->{xmlReader}, "start of readRow()");

  if ($self->{rowIndex} >= $self->{rowCount})
  {
    (defined($self->{xmlReader})) && $self->parkXmlReader();
    return undef;
  }

  # Handle repeated rows
  if (defined($self->{currentRow}) && ($self->{rowIndex} < ($self->{currentRow}->{rowDef}->{startIndex} + $self->{currentRow}->{rowDef}->{count})))
  {
    $self->{rowIndex}++;
    return $self->{currentRow}->{cellDefs};
  }
  
  # If no more rows, we're done reading
  #(!Spreadsheet::Nifty::XMLReaderUtils->atStartOfElement($self->{xmlReader}, 'table-row', $xmlns)) && return undef;
  (!Spreadsheet::Nifty::XMLReaderUtils->atStartOfElement($self->{xmlReader}, 'table-row', $xmlns)) && die("readRow(): Premature end of rows");

  # Read current row
  my $rowElement = $self->{xmlReader}->copyCurrentNode(1);
  Spreadsheet::Nifty::XMLReaderUtils->skipElement($self->{xmlReader});
  ($self->{xmlReader}->read() != 1) && die("readRow(): XML read error");

  # Decode rowDefs (each can span multiple rows)
  my $rowDef = Spreadsheet::Nifty::ODS::Decode->decodeRowDefinition($rowElement);
  $rowDef->{startIndex} = $self->{rowIndex};
  push(@{$self->{rowDefs}}, $rowDef);

  # Decode cellDefs (each can span multiple cells)
  my $cellDefs = [];
  my $children = [ $rowElement->getChildrenByTagName('*') ];
  for my $c (@{$children})
  {
    my $cellDef = Spreadsheet::Nifty::ODS::Decode->decodeCellDefinition($c);
    push(@{$cellDefs}, $cellDef);
  }

  # Remember current row
  $self->{currentRow} = {rowDef => $rowDef, cellDefs => $cellDefs};  

  $self->{rowIndex}++;

  #print main::Dumper($rowElement->toString(), $rowDef, $cellDefs);
  return $cellDefs;
}

sub seekRow($)
{
  my $self = shift();
  my ($rowIndex) = @_;

  ...;  # TODO
}

sub tellRow()
{
  my $self = shift();

  return $self->{rowIndex};
}

1;
