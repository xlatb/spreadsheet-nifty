#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::ODS::Cell;
use parent 'Spreadsheet::Nifty::Cell';

# === Class methods ===

sub new($$$;$)
{
  my $class = shift();
  my ($type, $value, $ctx, $private) = @_;

  my $self = $class->SUPER::new($type, $value);
  $self->{p} = $private // {};
  $self->{ctx} = sub { return $ctx; };  # Closure around context

  Scalar::Util::weaken($ctx);

  return $self;
}

# === Instance methods ===

sub stringValue()
{
  my $self = shift();

  if ($self->{t} == Spreadsheet::Nifty::TYPE_DATE)
  {
    if (defined($self->{p}->{value}))
    {
      return $self->{p}->{value};
    }
  }

  return $self->SUPER::stringValue();
}

1;
