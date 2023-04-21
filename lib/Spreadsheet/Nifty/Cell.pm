#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::Cell;

# === Class methods ===

sub new($$)
{
  my $class = shift();
  my ($type, $value) = @_;

  my $self = {};
  $self->{t} = $type;
  $self->{v} = $value;

  bless($self, $class);

  return $self;
}

# === Instance methods ===

sub dup()
{
  my $self = shift();

  my $new = {t => $self->{t}, v => $self->{v}};
  bless($new, ref($self));

  return $new;
}

sub value()
{
  my $self = shift();

  return $self->{v};
}

sub type()
{
  my $self = shift();

  return $self->{t};
}

sub formatString()
{
  ...;
}

sub formattedValue()
{
  ...;
}

1;
