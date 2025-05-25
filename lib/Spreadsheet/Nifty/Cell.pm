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

sub stringValue()
{
  my $self = shift();

  my $t = $self->{t};
  if ($t == Spreadsheet::Nifty::TYPE_NULL)
  {
    return undef;
  }
  elsif (($t == Spreadsheet::Nifty::TYPE_NUM) || ($t == Spreadsheet::Nifty::TYPE_STR))
  {
    return $self->{v};
  }
  elsif ($t == Spreadsheet::Nifty::TYPE_ERR)
  {
    return Spreadsheet::Nifty->errorName($self->{v});
  }
  elsif ($t == Spreadsheet::Nifty::TYPE_BOOL)
  {
    return ($self->{v} ? 'TRUE' : 'FALSE');
  }
  elsif ($t == Spreadsheet::Nifty::TYPE_DATE)
  {
    ...;
  }
  else
  {
    die('Unknown data type');
  }
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
