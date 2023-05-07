#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSB::Sheet;

use Spreadsheet::Nifty::XLSB::Cell;

# === Class methods ===

sub new()
{
  my $class = shift();
  my ($workbook, $sheetReader) = @_;

  my $self = {};
  $self->{workbook}    = $workbook;
  $self->{sheetReader} = $sheetReader;

  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub tellRow()
{
  my $self = shift();

  return $self->{sheetReader}->tellRow();
}

sub buildCell($)
{
  my $self = shift();
  my ($data) = @_;

  my $cell = Spreadsheet::Nifty::XLSB::Cell->new($data->{dataType}, $data->{value});
  return $cell;
}

sub readRow()
{
  my $self = shift();

  my $row = $self->{sheetReader}->readRow();
  (!defined($row)) && return;

  for (my $i = 0; $i < scalar(@{$row}); $i++)
  {
    (!defined($row->[$i])) && next;
    $row->[$i] = $self->buildCell($row->[$i]);
  }

  return $row;
}

sub getName()
{
  my $self = shift();

  return $self->{sheetReader}->getName();
}

sub getRowDimensions()
{
  my $self = shift();

  return $self->{sheetReader}->getRowDimensions();
}

sub getColDimensions()
{
  my $self = shift();

  return $self->{sheetReader}->getColDimensions();
}

1;
