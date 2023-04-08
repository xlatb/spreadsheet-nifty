#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLS::FileReader;

use IO::File;
use Encode qw();

use Spreadsheet::Nifty::CDF;
use Spreadsheet::Nifty::XLS;
use Spreadsheet::Nifty::XLS::BIFFReader;
use Spreadsheet::Nifty::XLS::RecordTypes;
use Spreadsheet::Nifty::XLS::SheetReader;
use Spreadsheet::Nifty::XLS::Sheet;
use Spreadsheet::Nifty::XLS::Decode;
use Spreadsheet::Nifty::XLS::Crypto;
use Spreadsheet::Nifty::XLS::Formula;

# === Class methods ===

sub new()
{
  my $class = shift();
  
  my $self = {};
  $self->{filename}   = undef;
  $self->{cdf}        = undef;
  $self->{dirEntries} = {};
  $self->{biff}       = undef;
  $self->{workbook}   = undef;
  $self->{strings}    = undef;
  $self->{offsets}    = {};
  $self->{debug}      = 0;
  $self->{sheetCount} = undef;
  $self->{password}   = undef;
#  $self->{worksheets}    = undef;
  
  bless($self, $class);

  return $self;
}

sub isFileSupported($)
{
  my $class = shift();
  my ($filename) = @_;

  my $file = IO::File->new();
  $file->open($filename, 'r') || die("open(): $!");

  my $cdf = Spreadsheet::Nifty::CDF->new($file);
  $cdf->open() || return 0;

  my $rootDir = $cdf->getRootDirectory();
  (defined($rootDir)) || return 0;

  my $workbook = $rootDir->getChildNamed('Workbook') // $rootDir->getChildNamed('Book');
  (defined($workbook)) || return 0;

  my $biff = Spreadsheet::Nifty::XLS::BIFFReader->new($workbook->open());
  my $rec = $biff->readRecord();
  (defined($rec)) || return 0;
  #(($rec->{name} // '') eq 'BOF') || return 0;
  ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::BOF) || return 0;

  return 1;
}

sub biffVersionName($)
{
  my $class = shift();
  my ($biffVersion) = @_;

  return {'0200' => 'BIFF2', '0300' => 'BIFF3', '0400' => 'BIFF4', '0500' => 'BIFF5', '0600' => 'BIFF8'}->{sprintf("%04X", $biffVersion)} // '(unknown)';
}

# === Instance methods ===

sub setPassword($)
{
  my $self = shift();
  my ($password) = @_;

  $self->{password} = $password;
  return;
}

sub setDebug($)
{
  my $self = shift();
  my ($debug) = @_;

  $self->{debug} = $debug;
  return;
}

sub open($)
{
  my $self = shift();
  my ($filename) = @_;

  my $file = IO::File->new();
  $file->open($filename, 'r') || return 0;

  my $cdf = Spreadsheet::Nifty::CDF->new($file);
  $cdf->open() || return 0;

  $self->{cdf} = $cdf;
  $self->{dirEntires} = {};
  $self->{dirEntries}->{root} = $cdf->getRootDirectory();

  return $self->read();
}

sub read()
{
  my $self = shift();

  if (defined(my $compObjStream = $self->{dirEntries}->{root}->getChildNamed("\x01CompObj")))
  {
    ($self->{debug}) && printf("read compObjStream\n");
    Spreadsheet::Nifty::XLS::Decode::decodeCompObjStream($compObjStream->getBytes());
  }

  my $workbookEntry = $self->{dirEntries}->{root}->getChildNamed("Workbook");
  if (!defined($workbookEntry))
  {
    $workbookEntry = $self->{dirEntries}->{root}->getChildNamed("Book");  # This name used in earlier BIFF5
  }
  (!defined($workbookEntry)) && return 0;
  $self->{dirEntries}->{workbook} = $workbookEntry;

  my $workbookHandle = $workbookEntry->open();
  (!defined($workbookHandle)) && return 0;

  $self->{biff} = Spreadsheet::Nifty::XLS::BIFFReader->new($workbookHandle);
  
  (!$self->readWorkbookStream()) && return 0;

  return 1;
}

sub readWorkbookStream($)
{
  my $self = shift();

  $self->{workbook} = {};
  $self->{workbook}->{fonts}       = [];
  $self->{workbook}->{formats}     = [];
  $self->{workbook}->{styles}      = [];
  $self->{workbook}->{boundsheets} = [];
  $self->{workbook}->{strings}     = [];

  $self->{offsets}->{workbook} = {};

  my $rec = $self->{biff}->readRecord();
  (!defined($rec)) && return 0;
  #($rec->{name} ne 'BOF') && return 0;
  ($rec->{type} != Spreadsheet::Nifty::XLS::RecordTypes::BOF) && return 0;
  $self->{offsets}->{workbook}->{bof} = $rec->{offset};

  my $bof = Spreadsheet::Nifty::XLS::Decode::decodeBOF($rec->{payload});
  ($bof->{type} != Spreadsheet::Nifty::XLS::BOF_TYPE_WORKBOOK) && return 0;
  $self->{workbook}->{bof} = $bof;

  my $biffVersion = $bof->{version};

  while (defined(my $rec = $self->{biff}->readRecord()))
  {
    #printf("Workbook rec: name '%s' length %d\n", $rec->{name}, $rec->{length});
    #if ($rec->{name} eq 'EOF')
    if ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::EOF)
    {
      $self->{offsets}->{workbook}->{eof} = $rec->{offset};
      $self->{offsets}->{workbook}->{next} = $self->{biff}->tell();
      last;
    }
    #elsif ($rec->{name} eq 'FilePass')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::FILE_PASS)
    {
      my $filePass = Spreadsheet::Nifty::XLS::Decode::decodeFilePass($rec->{payload}, $biffVersion);
      $self->{workbook}->{encryption} = $filePass;

      if ($filePass->{type} == 0)
      {
        # Xor obfuscation
        my $xor = Spreadsheet::Nifty::XLS::Crypto->xor($self->{password} // $Spreadsheet::Nifty::XLS::defaultPassword);
        if (!$xor->checkPassword($self->{workbook}->{encryption}->{verifier}))
        {
          die("Encrypted file (xor), unknown password");
        }
        
        $self->{biff}->setDecryptor($xor);
      }
      elsif (($filePass->{type} == 1) && ($filePass->{version}->{major} == 1) && ($filePass->{version}->{minor} == 1))
      {
        # RC4 encryption (non-CryptoAPI)
        my $rc4 = Spreadsheet::Nifty::XLS::Crypto->rc4($self->{password} // $Spreadsheet::Nifty::XLS::defaultPassword, $self->{workbook}->{encryption}->{salt});
        if (!$rc4->checkPassword($self->{workbook}->{encryption}->{verifier}, $self->{workbook}->{encryption}->{verifierHash}))
        {
          die("Encrypted file (RC4), unknown password");
        }
        
        $self->{biff}->setDecryptor($rc4);
      }
      elsif (($filePass->{type} == 1) && (($filePass->{version}->{major} >= 2) || ($filePass->{version}->{major} >= 4)) && ($filePass->{version}->{minor} == 2))
      {
        # CryptoAPI
        if ($filePass->{header}->{encryptionType} == 0x6801)  # RC4
        {
          my $c = Spreadsheet::Nifty::XLS::Crypto->cryptoApiRC4($self->{password} // $Spreadsheet::Nifty::XLS::defaultPassword, $filePass->{header}, $filePass->{verifier});
          (!$c->checkPassword()) && die("Encrypted file (CryptoAPI RC4), unknown password");
          $self->{biff}->setDecryptor($c);
        }
        else
        {
          ...;
        }
      }
      else
      {
        die("Unhandled encryption type");
      }
    }
    #elsif ($rec->{name} eq 'InterfaceHdr')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::INTERFACE_HDR)
    {
      $self->readInterface($rec);
    }
    #elsif ($rec->{name} eq 'RRTabId')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::RR_TAB_ID)
    {
      my $decoder = StructDecoder->new($rec->{payload});
      $self->{workbook}->{tabIds} = $decoder->decodeArray('u16', $decoder->bytesLeft() / 2);
    }
    #elsif ($rec->{name} eq 'Protect')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::PROTECT)
    {
      $self->{workbook}->{protect} = unpack('v', $rec->{payload});
    }
    #elsif ($rec->{name} eq 'Date1904')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::DATE_1904)
    {
      $self->{workbook}->{flagDate1904} = unpack('v', $rec->{payload});
    }
    #elsif ($rec->{name} eq 'Font')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::FONT)
    {
      push(@{$self->{workbook}->{fonts}}, Spreadsheet::Nifty::XLS::Decode::decodeFont($rec->{payload}, $biffVersion));
    }
    #elsif ($rec->{name} eq 'Format')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::FORMAT)
    {
      push(@{$self->{workbook}->{formats}}, Spreadsheet::Nifty::XLS::Decode::decodeFormat($rec->{payload}, $biffVersion));
    }
    #elsif ($rec->{name} eq 'DXF')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::DXF)
    {
      # TODO: Differential formats
    }
    #elsif ($rec->{name} eq 'XF')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::XF)
    {
      push(@{$self->{workbook}->{xfs}}, Spreadsheet::Nifty::XLS::Decode::decodeXF($rec->{payload}, $biffVersion));
    }
    #elsif ($rec->{name} eq 'XFCRC')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::XF_CRC)
    {
      # TODO: XF CRC
    }
    #elsif ($rec->{name} eq 'XFExt')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::XF_EXT)
    {
      # TODO: FXExt
    }
    #elsif ($rec->{name} eq 'Style')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::STYLE)
    {
      push(@{$self->{workbook}->{styles}}, Spreadsheet::Nifty::XLS::Decode::decodeStyle($rec->{payload}, $biffVersion));
    }
    #elsif ($rec->{name} eq 'Palette')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::PALETTE)
    {
      $self->{workbook}->{palette} = Spreadsheet::Nifty::XLS::Decode::decodePalette($rec->{payload});
    }
    #elsif ($rec->{name} eq 'BoundSheet8')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::BOUND_SHEET_8)
    {
      push(@{$self->{workbook}->{boundsheets}}, Spreadsheet::Nifty::XLS::Decode::decodeBoundSheet8($rec->{payload}, $biffVersion));
    }
    #elsif ($rec->{name} eq 'Lbl')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::LBL)
    {
      push(@{$self->{workbook}->{labels}}, Spreadsheet::Nifty::XLS::Decode::decodeLbl($rec->{payload}, $biffVersion));
    }
    #elsif ($rec->{name} eq 'SST')
    elsif ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::SST)
    {
      $self->readStringTable($rec->{payload});
    }
  }

  return 1;
}

# Given a sheet index, returns the offset of that sheet's BOF record within the file.
sub findSheetBOF($)
{
  my $self = shift();
  my ($index) = @_;

  # If offset already known, just return it
  my $offset = $self->{offsets}->{sheets}->[$index]->{bof};
  (defined($offset)) && return $offset;

  # If the workbook has a BoundSheet8 record for the stream, try that offset
  $offset = $self->findSheetBOFViaBoundSheet($index);
  (defined($offset)) && return $offset;

  # Fall back to scanning
  my $offsets = $self->scanSheetOffsets($index);
  (defined($offsets)) && return $offsets->{bof};

  return undef;
}

sub findSheetBOFViaBoundSheet($)
{
  my $self = shift();
  my ($index) = @_;

  my $boundsheet = $self->{workbook}->{boundsheets}->[$index];
  (!defined($boundsheet)) && return undef;  # No BoundSheet8

  my $offset = $boundsheet->{offset};
  #printf("  Boundsheet offset: %d\n", $offset);
  ($offset == 0) && return undef;  # Missing offset

  $self->{biff}->seek($offset);
  my $rec = $self->{biff}->readRecord();
  (!defined($rec)) && return undef;  # No record found
  ($rec->{type} ne Spreadsheet::Nifty::XLS::RecordTypes::BOF) && return undef;  # Not a BOF record

  #$self->{offsets}->{sheets}->[$index]->{bof} = $offset;
  $self->recordSheetOffsets($index, {bof => $offset});

  return $offset;
}

# Given a sheet index, scans the file to find the sheet substream offsets.
# Returns a hash of the offsets if found.
sub scanSheetOffsets($)
{
  my $self = shift();
  my ($index) = @_;

  (!defined($self->{offsets}->{workbook}->{next})) && die("scanSheetOffsets(): No known end of workbook stream");

  # All offsets already known? Just return them
  if (defined($self->{offsets}->{sheets}->[$index]))
  {
    my $offsets = $self->{offsets}->{sheets}->[$index];
    if (defined($offsets->{bof}) && defined($offsets->{eof}) && defined($offsets->{next}))
    {
      return $offsets;
    }
  }

  # Use the highest-indexed preceding sheet with an offset as our starting point
  my ($offset, $s);
  for (my $i = $index; $i >= 0; $i--)
  {
    my $o = $self->{offsets}->{sheets}->[$i]->{bof};
    if (defined($o))
    {
      $s = $i;
      $offset = $o;
      last;
    }
  }

  # If we didn't find a preceding sheet, start at the end of the workbook
  if (!defined($offset))
  {
    $s = 0;
    $offset = $self->{offsets}->{workbook}->{next};
  }

  while ($s <= $index)
  {
    # Read the BOF.
    $self->{biff}->seek($offset);
    my $rec = $self->{biff}->readRecord();
    #(!defined($rec) || ($rec->{name} ne 'BOF')) && return 0;
    (!defined($rec) || ($rec->{type} != Spreadsheet::Nifty::XLS::RecordTypes::BOF)) && return undef;
    if (!defined($self->{offsets}->{sheets}->[$s]->{bof}))
    {
      $self->recordSheetOffsets($s, {bof => $rec->{offset}});
      #$self->{offsets}->{sheets}->[$s]->{bof} = $rec->{offset};
    }

    # Scan until EOF
    while (defined($rec = $self->{biff}->readRecord()))
    {
      if ($rec->{name} eq 'EOF')
      {
        $self->recordSheetOffsets($s, {eof => $rec->{offset}, next => $self->{biff}->tell()});
        #$self->{offsets}->{sheets}->[$s]->{eof}  = $rec->{offset};
        #$self->{offsets}->{sheets}->[$s]->{next} = $self->{biff}->tell();
        last;
      }
    }

    $s++;
  }

  return $self->{offsets}->{sheets}->[$index];
}

sub getSheetOffset($$)
{
  my $self = shift();
  my ($index, $name) = @_;

  return $self->{offsets}->{sheets}->[$index]->{$name};
}

sub recordSheetOffsets($$)
{
  my $self = shift();
  my ($index, $offsets) = @_;

  (!defined($self->{offsets}->{sheets})) && do { $self->{offsets}->{sheets} = []; };
  (!defined($self->{offsets}->{sheets}->[$index])) && do { $self->{offsets}->{sheets}->[$index] = {}; };

  for my $k (keys(%{$offsets}))
  {
    my $new = $offsets->{$k};
    (!defined($new)) && next;

    my $old = $self->{offsets}->{sheets}->[$index]->{$k};
    if (defined($old) && ($old != $new))
    {
      warn(sprintf("Sheet %d offset '%s' changed %s -> %s", $index, $k, $old, $new));
    }

    $self->{offsets}->{sheets}->[$index]->{$k} = $new;
  }

  return;  
}

sub readInterface($$)
{
  my $self = shift();
  my ($hdr) = @_;

  # Ignore records until InterfaceEnd
  while (defined(my $rec = $self->{biff}->readRecord()))
  {
    #(($rec->{name} // '') eq 'InterfaceEnd') && last;
    ($rec->{type} == Spreadsheet::Nifty::XLS::RecordTypes::INTERFACE_END) && last;
  }

  return;
}

# NOTE: These can be broken across records, in which case they require special
#  handling of following Continue records.
sub readXLUnicodeRichExtendedString($)
{
  my $self = shift();
  my ($decoder) = @_;

  # Reads a following Continue record into the decoder input buffer if there
  #  are fewer than the given number of bytes remaining. Not for use inside
  #  of strings, which have their own special Continue handling.
  my $readContinueIfNeeded = sub
  {
    my ($wantSize) = @_;

    my $remainSize = $decoder->bytesLeft();
    while ($wantSize > $remainSize)
    {
      #printf("Reading a Continue\n");
      my $rec = $self->{biff}->readRecord();
      (!defined($rec)) && die("decodeXLUnicodeRichExtendedString(): Expected a following Continue record, got EOF");
      ($rec->{type} != Spreadsheet::Nifty::XLS::RecordTypes::CONTINUE) && die("decodeXLUnicodeRichExtendedString(): Expected a following Continue record, got record type $rec->{type}");
      $decoder->appendBytes($rec->{payload});
      $remainSize = $decoder->bytesLeft();
    }
  };

  my $data = $decoder->decodeHash(['cch:u16', 'flags:u8']);

  if ($data->{flags} & 0x08)
  {
    $data->{cRun} = $decoder->decodeField('u16');
  }

  if ($data->{flags} & 0x04)
  {
    $data->{cbExtRst} = $decoder->decodeField('u32');
  }

  # Reading the string is complicated by the fact that:
  # 1. If the string is longer than the payload, its remainder is in a following Continue record
  # 2. The Continue record can only happen during a variable-length field
  # 3. The Continue record begins with a byte containing a flag for whether the string is UCS-2 or Latin-1
  # So the string can contain multiple parts, in different encodings, that we need to concatenate
  my $str = '';
  my $charsLeft = $data->{cch};
  my $ucs2 = $data->{flags} & 0x01;
  while ($charsLeft)
  {
    my $wantSize = ($ucs2) ? ($charsLeft * 2) : $charsLeft;
    my $remainSize = $decoder->bytesLeft();
    my $size = ($wantSize <= $remainSize) ? $wantSize : $remainSize;
    #printf("Wantsize %d remainsize %d\n", $wantSize, $remainSize);

    if ($ucs2)
    {
      my $chunk = Encode::decode('UTF-16LE', $decoder->getBytes($size));
      #printf("UCS-2 CHUNK: %s\n", $chunk);
      $str .= $chunk;
      $charsLeft -= ($size >> 1);
    }
    else
    {
      my $chunk = Encode::decode('iso-8859-1', $decoder->getBytes($size));
      #printf("Latin-1 CHUNK: %s\n", $chunk);
      $str .= $chunk;
      $charsLeft -= $size;
    }

    if ($charsLeft)
    {
      #printf("Time for a Continue section!\n");
      my $rec = $self->{biff}->readRecord();
      (!defined($rec)) && die("decodeXLUnicodeRichExtendedString(): Expected a following Continue record, got EOF");
      ($rec->{type} != 0x003C) && die("decodeXLUnicodeRichExtendedString(): Expected a following Continue record, got record type $rec->{type}");
      $decoder->setBytes($rec->{payload});
      $ucs2 = $decoder->decodeField('u8');  # Read new UCS-2 flag
      #printf("UCS2 value: %d\n", $ucs2);
    }
  }

  $data->{str} = $str;

  if ($data->{flags} & 0x08)
  {
    $readContinueIfNeeded->(4 * $data->{cRun});
#    my $wantSize = 4 * $data->{cRun};
#    my $remainSize = $decoder->bytesLeft();
#    if ($wantSize > $remainSize)
#    {
#      #printf("Time for a Continue section! (in rgRun)\n");
#      my $rec = $biff->readRecord();
#      (!defined($rec)) && die("decodeXLUnicodeRichExtendedString(): Expected a following Continue record, got EOF");
#      ($rec->{type} != 0x003C) && die("decodeXLUnicodeRichExtendedString(): Expected a following Continue record, got record type $rec->{type}");
#      $decoder->appendBytes($rec->{payload});
#    }
    $data->{rgRun} = $decoder->decodeArray(['ich:u16', 'ifnt:u16'], $data->{cRun});
  }

  if ($data->{flags} & 0x04)
  {
    $data->{ExtRst} = $decoder->getBytes($data->{cbExtRst});  # Skip over ExtRst
  }

  return $data;  
}

sub readStringTable($$)
{
  my $self = shift();
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $sst = $decoder->decodeHash(['refCount:u32', 'strCount:u32']);

  my $strings = [];
  while (scalar(@{$strings}) < $sst->{strCount})
  {
    while ($decoder->bytesLeft())
    {
      my $s = $self->readXLUnicodeRichExtendedString($decoder);
      push(@{$strings}, $s);
    }

    my $rec = $self->{biff}->readRecordIfType(Spreadsheet::Nifty::XLS::RecordTypes::CONTINUE);
    (!defined($rec)) && last;
    $decoder->setBytes($rec->{payload});
  }

  # TODO: Check if total equals the number of strings read?

  $self->{workbook}->{strings} = $strings;
  return;
}

sub getLabel($)
{
  my $self = shift();
  my ($index) = @_;

  return $self->{workbook}->{labels}->[$index];
}

sub getSheetNames()
{
  my $self = shift();

  return [ map({ $_->{name} } @{$self->{workbook}->{boundsheets}}) ];
}

sub getSheetCount()
{
  my $self = shift();

  (defined($self->{sheetCount})) && do return $self->{sheetCount};

  # If we have a nonzero number of BoundSheet8 records, use that count
  if (defined($self->{workbook}->{boundsheets}))
  {
     my $count = scalar(@{$self->{workbook}->{boundsheets}});
     if ($count > 0)
     {
       $self->{sheetCount} = $count;
       return $count;
     }
  }

  # Fallback - Scan the file until we run out of sheets
  my $count = 0;
  while ($self->scanSheetOffsets($count))
  {
    $count++;
  }

  $self->{sheetCount} = $count;
  return $count;
}

sub openSheet($)
{
  my $self = shift();
  my ($index) = @_;

  # Bounds check
  my $count = $self->getSheetCount();
  (($index < 0) || ($index >= $count)) && return undef;  # Out of bounds

  # Find sheet offset
  my $offset = $self->findSheetBOF($index);
  (!defined($offset)) && return undef;  # Not found

  # Open a new independent I/O handle so that the sheet can seek independently.
  my $io = $self->{dirEntries}->{workbook}->open();
  (!defined($io)) && return undef;  # Could not open

  # Create a new BIFF record reader but copy the old decryption settings, if any
  my $biff = Spreadsheet::Nifty::XLS::BIFFReader->new($io);
  $biff->copyDecryptor($self->{biff});

  # Create a SheetReader for the sheet
  my $sheetReader = Spreadsheet::Nifty::XLS::SheetReader->new($self, $index, $offset, $biff);
  (!$sheetReader->open()) && return undef;

  # Wrap the SheetReader in a Sheet
  my $sheet = Spreadsheet::Nifty::XLS::Sheet->new($self, $index, $sheetReader);
  return $sheet;
}

1;
