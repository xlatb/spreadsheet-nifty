#!/usr/bin/perl -w
use strict;
use warnings;

package Spreadsheet::Nifty::IndexedColors;

# Legacy XLS palette of 56 colours.
# The colours in the palette can be customized on a per-file basis, but this
#  is the default palette.
my $legacyColors =
[
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'
];

# === Class methods ===

sub new()
{
  my $class = shift();

  my $self = {};
  $self->{colors} = undef;  # Undef means fall back to legacy colours

  bless($self, $class);
  return $self;
}

sub getLegacyColor($)
{
  my $class = shift();
  my ($index) = @_;

  ($index < 0) && return undef;
  ($index > 65) && return undef;

  ($index == 64) && return '000000';  # System foreground
  ($index == 65) && return 'FFFFFF';  # System background

  # First 8 colours are copies of the second 8 for backwards compatibility reasons.
  ($index < 8) && return $legacyColors->[$index];

  return $legacyColors->[$index - 8];
}

# === Instance methods ===

sub addColor($)
{
  my $self = shift();
  my ($c) = @_;

  (!defined($self->{colors})) && do { $self->{colors} = []; };

  push(@{$self->{colors}}, $c);

  return;
}

sub getColor($)
{
  my $self = shift();
  my ($i) = @_;

  (!defined($self->{colors})) && return $self->getLegacyColor($i);
  
  return $self->{colors}->[$i];
}

1;
