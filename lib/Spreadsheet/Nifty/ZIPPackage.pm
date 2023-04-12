#!/usr/bin/perl -w
use warnings;
use strict;

use Spreadsheet::Nifty::ZIPReader;

package Spreadsheet::Nifty::ZIPPackage;

our $namespaces =
{
  relationshipDefs => 'http://schemas.openxmlformats.org/package/2006/relationships',
  relationshipRefs => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',

  officeDocument => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  sharedStrings => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
  styles => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
};

use Archive::Zip;
use XML::LibXML qw(:libxml);
use XML::LibXML::Reader;

# === Class methods ===

sub new()
{
  my $class = shift();
  
  my $self = {};
  $self->{zip}      = undef;
  $self->{filename} = undef;
  
  bless($self, $class);

  $self->{zip} = Archive::Zip->new();
  
  return $self;
}

# Given an initial path and a target path, resolves to an absolute path within the ZIP.
sub resolveZipPath($$)
{
  my $class = shift();
  my ($initial, $target) = @_;

  my $path;
  if ($target =~ m#^/#)
  {
    # Target starts with a slash, so it replaces any existing path
    $path = ($target =~ s#^/##r);
  }
  else
  {
    # Target does not start with slash, so append to existing path
    $path = $initial . '/' . $target;
  }

  # Runs of slashes become single slashes
  $path =~ s#//+#/#g;

  # Handle '..' elements
  my $elements = [ split('/', $path) ];
  my $e = 1;
  while ($e < scalar(@{$elements}))
  {
    if ($elements->[$e] eq '..')
    {
      splice(@{$elements}, $e - 1, 2);
    }
    else
    {
      $e++;
    }
  }
  $path = join('/', @{$elements});

  return $path;
}

# 'a/b/c.xml' => 'a/b/_rels/c.xml.rels'
sub partnameToRelationshipsPart($)
{
  my $class = shift();
  my ($partname) = @_;

  my $elements = [ split('/', $partname) ];
  my $final = pop(@{$elements});

  my $relpartname = join('/', @{$elements}) . '/_rels/' . $final . '.rels';
  #printf("partnameToRelationshipsPart %s -> %s\n", $partname, $relpartname);
  return $relpartname;
}

sub partnameAncestor($$)
{
  my $class = shift();
  my ($partname, $steps) = @_;

  my $elements = [ split('/', $partname) ];
  for (my $i = 0; $i < $steps; $i++)
  {
    pop(@{$elements});
  }

  return join('/', @{$elements});
}

# Returns the first relationship of the given type, or undef if none.
sub getRelationshipByType($$)
{
  my $class = shift();
  my ($rels, $type) = @_;

  for my $id (keys(%{$rels}))
  {
    ($rels->{$id}->{type} eq $type) && return $rels->{$id};
  }

  return undef;
}

# === Instance methods ===

sub open($)
{
  my $self = shift();
  my ($filename) = @_;

  $self->{filename} = $filename;
  
  my $result = $self->{zip}->read($filename);
  ($result != Archive::Zip::AZ_OK) && return 0;

  return 1;
}

sub openMember($)
{
  my $self = shift();
  my ($membername) = @_;

  # Initial slash is removed
  $membername =~ s#^/##;

  my $zipMember = $self->{zip}->memberNamed($membername);
  (!$zipMember) && die("Member '${membername}' not found in this ZIP package\n");

  my $zipReader = Spreadsheet::Nifty::ZIPReader->new($zipMember);
  return $zipReader;
}

sub readRelationshipsMember($)
{
  my $self = shift();
  my ($membername) = @_;

  my $xmlns = $namespaces->{relationshipDefs};

  my $relationships = {};

  my $zipReader = $self->openMember($membername);

  my $xmlReader = XML::LibXML::Reader->new({IO => $zipReader});

  my $cwd = $self->partnameAncestor($membername, 2);  # These partnames normally end in /_rels/x.rels

  my $status;
  while (($status = $xmlReader->nextElement('Relationship', $xmlns)) == 1)
  {
    my $node = $xmlReader->copyCurrentNode(1);

    my $id = $node->getAttribute('Id');
    my $type = $node->getAttribute('Type');
    my $target = $node->getAttribute('Target');
    my $partname = $self->resolveZipPath($cwd, $target);  # Convert to absolute path

    ((!defined($id)) || (!defined($type)) || (!defined($target))) && next;

    $relationships->{$id} = {id => $id, type => $type, target => $target, partname => $partname};
  }

  return $relationships;
}

1;
