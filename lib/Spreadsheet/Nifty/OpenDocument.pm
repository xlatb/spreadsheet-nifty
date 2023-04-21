#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::OpenDocument;

our $namespaces =
{
  'manifest' => 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
  'dsig'     => 'urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0',
  'pkg'      => 'http://docs.oasis-open.org/ns/office/1.2/meta/pkg#',
  'ds'       => 'http://www.w3.org/2000/09/xmldsig#',
};

use Archive::Zip;
use XML::LibXML qw(:libxml);

use Spreadsheet::Nifty::ZIPReader;

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

sub hasMember($)
{
  my $self = shift();
  my ($membername) = @_;

  $membername =~ s#^/##;  # Initial slash is removed, if any

  my $zipMember = $self->{zip}->memberNamed($membername);
  return defined($zipMember);
}

sub openMember($)
{
  my $self = shift();
  my ($membername) = @_;

  $membername =~ s#^/##;  # Initial slash is removed, if any

  my $zipMember = $self->{zip}->memberNamed($membername);
  (!$zipMember) && die("Member '${membername}' not found in this ZIP package\n");

  my $zipReader = Spreadsheet::Nifty::ZIPReader->new($zipMember);
  return $zipReader;
}

sub readMimetype()
{
  my $self = shift();

  my $mimetype = $self->{zip}->contents('mimetype');

  return $mimetype;
}

sub readManifest()
{
  my $self = shift();

  (!$self->hasMember('META-INF/manifest.xml')) && return undef;

  my $zipReader = $self->openMember('META-INF/manifest.xml');

  my $xmlns = $namespaces->{manifest};

  my $manifest = {};

  my $xmlParser = XML::LibXML->new({load_ext_dtd => 0, validation => 0});
  my $doc = $xmlParser->load_xml({IO => $zipReader});

  my $root = $doc->documentElement();
  if (($root->namespaceURI() ne $xmlns) || ($root->localName() ne 'manifest'))
  {
    return undef;  # Root element is not <manifest:manifest/>
  }

  $manifest->{version} = $root->getAttributeNS($xmlns, 'version');

  # Read file entries
  $manifest->{entries} = [];
  for my $kid ($root->getChildrenByTagNameNS($xmlns, 'file-entry'))
  {
    my $entry = {};
    $entry->{path} = $kid->getAttributeNS($xmlns, 'full-path');
    $entry->{mimetype} = $kid->getAttributeNS($xmlns, 'media-type');
    $entry->{version}  = $kid->getAttributeNS($xmlns, 'version');
    $entry->{size} = $kid->getAttributeNS($xmlns, 'size');
    $entry->{viewMode} = $kid->getAttributeNS($xmlns, 'preferred-view-mode');

    my ($encryption) = $kid->getChildrenByTagNameNS($xmlns, 'encryption-data');
    if (defined($encryption))
    {
      $entry->{encryption} = {};
      $entry->{encryption}->{checksum} = $encryption->getAttributeNS($xmlns, 'checksum');
      $entry->{encryption}->{checksumType} = $encryption->getAttributeNS($xmlns, 'checksum-type');

      # TODO: Child elements <manifest:algorithm/>, <manifest:key-derivation/>, <manifest:start-key-generation/>
    }

    push(@{$manifest->{entries}}, $entry);
  }


  return $manifest;
}

1;
