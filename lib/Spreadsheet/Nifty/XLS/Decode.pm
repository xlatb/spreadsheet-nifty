#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLS::Decode;

use Spreadsheet::Nifty::Utils;

use Encode qw();

# === Utilities ===

# An Rk is a number packed into a 32-bit value in a goofy way.
sub translateRk($)
{
  my ($value) = @_;

  my $flagA = $value & 0x01;
  my $flagB = $value & 0x02;
  $value &= 0xFFFFFFFC;  # Discard flag bits

  #printf("  Intermediate: 0x%08X\n", $value);
  if ($flagB)
  {
    # Signed integer
    $value >>= 2;
    #printf("  Signed integer: %d\n", $value);
  }
  else
  {
    # This worked but is unlikely to be portable
    #my $packed = pack('C8', reverse(unpack('C8', pack('N', $value) . pack('V', 0))));
    #$value = unpack('d', $packed');
    my $sign     = ($value & 0x80000000) >> 31;
    my $exponent = ($value & 0x7FF00000) >> 20;
    my $mantissa = ($value & 0x000FFFFF) << 32;
    #printf("  parts: sign %d exponent %d mantissa %d\n", $sign, $exponent, $mantissa);
    $value = Spreadsheet::Nifty::Utils->ieeePartsToValue($sign, $exponent, $mantissa, 11, 52, 1023);
  }
  if ($flagA)
  {
    #printf("  Divide by 100\n");
    $value = $value / 100;
  }
  #printf("  Final value: %s\n", $value);

  return $value;
}

# === Single-field decoders ===

sub decodeNoLenAnsiString($$)
{
  my $decoder = shift();
  my ($length) = @_;

  my $bytes = $decoder->getBytes($length);
  return Encode::decode('cp-1252', $bytes);
}

sub decodeLen8AnsiString($)
{
  my $decoder = shift();

  my $length = $decoder->decodeField('u8');
  my $bytes = $decoder->getBytes($length);
  return Encode::decode('cp-1252', $bytes);
}

sub decodeLen16AnsiString($)
{
  my $decoder = shift();

  my $length = $decoder->decodeField('u16');
  my $bytes = $decoder->getBytes($length);
  return Encode::decode('cp-1252', $bytes);
}

sub decodeLen32AnsiString($)
{
  my $decoder = shift();

  my $length = $decoder->decodeField('u32');
  my $bytes = $decoder->getBytes($length);
  return Encode::decode('cp-1252', $bytes);
}

# String with no character count. Character count must be known.
# Also known as "XLUnicodeStringNoCch".
sub decodeNoLenXLUnicodeString($$)
{
  my $decoder = shift();
  my ($count) = @_;

  my $flags = unpack('C', $decoder->getBytes(1));
  my $size = ($flags & 0x01) ? ($count * 2) : $count;

  my $payload = $decoder->getBytes($size);

  my $str;
  if ($flags & 0x01)
  {
    $str = Encode::decode('UTF-16LE', $payload);
  }
  else
  {
    $str = Encode::decode('iso-8859-1', $payload);
  }

  return $str;
}

sub decodeShortXLUnicodeString($)
{
  my $decoder = shift();

  my $charCount = $decoder->decodeField('u8');
  my $flags     = $decoder->decodeField('u8');

  my $str;
  if ($flags & 0x01)
  {
    my $bytes = $decoder->getBytes($charCount * 2);
    $str = Encode::decode('UTF-16LE', $bytes);
  }
  else
  {
    my $bytes = $decoder->getBytes($charCount);
    $str = Encode::decode('iso-8859-1', $bytes);
  }

  return $str;
}

sub decodeXLUnicodeString($)
{
  my $decoder = shift();

  my $charCount = $decoder->decodeField('u16');
  my $flags     = $decoder->decodeField('u8');
  
  my $str;
  if ($flags & 0x01)
  {
    my $bytes = $decoder->getBytes($charCount * 2);
    $str = Encode::decode('UTF-16LE', $bytes);
  }
  else
  {
    my $bytes = $decoder->getBytes($charCount);
    $str = Encode::decode('iso-8859-1', $bytes);
  }

  return $str;
}

sub decodeNullTerminatedUTF16LEString($)
{
  my $decoder = shift();

  my $bytes;
  while (1)
  {
    my $char = $decoder->decodeField('bytes[2]');
    ($char eq "\x00\x00") && last;  # Null terminator
    $bytes .= $char;
  }

  my $str = Encode::decode('UTF-16LE', $bytes);
  return $str;
}

# === Structure decoders ===

sub decodeCompObjStream($)
{
  my ($bytes) = @_;
  
  my $decoder = StructDecoder->new($bytes);
  $decoder->registerType('len32AnsiString', \&decodeLen32AnsiString);
  $decoder->skipBytes(28);
  
  my $compObj = {};
  $compObj->{userType} = $decoder->decodeField('len32AnsiString');

  return $compObj;
}

sub decodeBOF($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $bof = $decoder->decodeHash(['version:u16', 'type:u16', 'build:u16', 'year:u16']);

  if (($bof->{version} == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF8) && ($decoder->bytesLeft() == 8))
  {
    $bof->{bits1} = $decoder->decodeField('u32');
    $bof->{bits2} = $decoder->decodeField('u32');
  }

  return $bof;
}

sub decodeFilePass($$)
{
  my ($bytes, $version) = @_;

  my $decoder = StructDecoder->new($bytes);

  my $filepass = {};

  if ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF5)
  {
    $filepass->{type} = 0;  # Always XOR obfuscation
    $filepass->{key}      = $decoder->decodeField('u16');
    $filepass->{verifier} = $decoder->decodeField('u16');
    return $filepass;
  }

  $filepass->{type} = $decoder->decodeField('u16');

  if ($filepass->{type} == 0)
  {
    # XOR obfuscation
    $filepass->{key}      = $decoder->decodeField('u16');
    $filepass->{verifier} = $decoder->decodeField('u16');
  }
  elsif ($filepass->{type} == 1)
  {
    $filepass->{version} = $decoder->decodeHash(['major:u16', 'minor:u16']);
    if ($filepass->{version}->{major} == 1)
    {
      $filepass->{salt}         = $decoder->getBytes(16);
      $filepass->{verifier}     = $decoder->getBytes(16);
      $filepass->{verifierHash} = $decoder->getBytes(16);
    }
    elsif ((($filepass->{version}->{major} >= 2) || ($filepass->{version}->{major} <= 4)) && ($filepass->{version}->{minor} == 2))
    {
      # CryptoAPI encryption
      $filepass->{flags}      = $decoder->decodeField('u32');
      $filepass->{headerSize} = $decoder->decodeField('u32');
      $filepass->{header}     = decodeCryptoAPIHeader($decoder->getBytes($filepass->{headerSize}));
      $filepass->{verifier}   = decodeCryptoAPIVerifier($decoder->getBytes($decoder->bytesLeft()));
    }
    else
    {
      ...;
    }
  }
  else
  {
    ...;
  }

  return $filepass;
}

sub decodeCryptoAPIHeader($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  $decoder->registerType('nullTerminatedUTF16LEString', \&decodeNullTerminatedUTF16LEString);
  my $header = $decoder->decodeHash(['flags:u32', ':u32', 'encryptionType:u32', 'hashType:u32', 'keySize:u32', 'providerType:u32', ':u32', ':u32', 'providerName:nullTerminatedUTF16LEString']);
  return $header;
}

sub decodeCryptoAPIVerifier($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $verifier = $decoder->decodeHash(['saltSize:u32', 'salt:bytes[saltSize]', 'encryptedVerifier:bytes[16]', 'verifierHashSize:u32', 'encryptedVerifierHash:bytes[verifierHashSize]']);
  return $verifier;
}

sub decodeFont($$)
{
  my ($bytes, $version) = @_;

  my $decoder = StructDecoder->new($bytes);
  $decoder->registerType('ShortXLUnicodeString', \&decodeShortXLUnicodeString);
  $decoder->registerType('len8AnsiString', \&decodeLen8AnsiString);

  my $nameType = ($version <= Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF5) ? 'len8AnsiString' : 'ShortXLUnicodeString';
  my $font = $decoder->decodeHash(['height:u16', 'flags:u16', 'color:u16', 'weight:u16', 'subOrSuper:u16', 'underline:u8', 'family:u8', 'charset:u8', ':u8', "font:${nameType}"]);

  return $font;
}

sub decodeFormat($$)
{
  my ($bytes, $version) = @_;

  my $decoder = StructDecoder->new($bytes);
  $decoder->registerType('XLUnicodeString', \&decodeXLUnicodeString);
  $decoder->registerType('len8AnsiString', \&decodeLen8AnsiString);

  my $stringType = ($version <= Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF5) ? 'len8AnsiString' : 'XLUnicodeString';

  my $format = $decoder->decodeHash(['id:u16', "format:${stringType}"]);
  return $format;
}

sub decodeXF($$)
{
  my ($bytes, $version) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $xf = $decoder->decodeHash(['font:u16', 'format:u16', 'flags:u16']);
  $xf->{parent} = $xf->{flags} & 0x0FFF;
  $xf->{version} = $version;

  $xf->{align} = $decoder->decodeField('u8');

  if ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF5)
  {
    $xf->{orientation} = $decoder->decodeField('u8');
    $xf->{styling}     = $decoder->getBytes(8);
  }
  else
  {
    $xf->{orientation} = $decoder->decodeField('u8');
    $xf->{direction}   = $decoder->decodeField('u8');
    $xf->{groups}      = $decoder->decodeField('u8');
    $xf->{styling}     = $decoder->getBytes(10);
  }

  return $xf;
}

sub decodeStyle($$)
{
  my ($bytes, $version) = @_;

  my $decoder = StructDecoder->new($bytes);
  $decoder->registerType('ShortXLUnicodeString', \&decodeXLUnicodeString);
  $decoder->registerType('len8AnsiString', \&decodeLen8AnsiString);

  my $xf = $decoder->decodeField('u16');

  my $style = {};
  $style->{xf}    = $xf & 0x0FFF;
  $style->{flags} = $xf & 0xF000;

  if ($style->{flags} & 0x8000)
  {
    $style->{builtin} = $decoder->decodeField('u16');
  }

  if ($decoder->bytesLeft())
  {
    my $nameType = ($version <= Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF5) ? 'len8AnsiString' : 'ShortXLUnicodeString';
    $style->{name} = $decoder->decodeField($nameType);
  }

  return $style;
}

sub decodePalette($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);

  my $count = $decoder->decodeField('u16');
  my $palette = $decoder->decodeArray('u32', $count);
  return $palette;
}

sub decodeBoundSheet8($$)
{
  my ($bytes, $version) = @_;

  my $decoder = StructDecoder->new($bytes);
  $decoder->registerType('len8AnsiString', \&decodeLen8AnsiString);
  $decoder->registerType('ShortXLUnicodeString', \&decodeShortXLUnicodeString);

  my $nameType = ($version <= Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF5) ? 'len8AnsiString' : 'ShortXLUnicodeString';

  my $boundsheet = $decoder->decodeHash(['offset:u32', 'hidden:u8', 'type:u8', "name:${nameType}"]);
  return $boundsheet;
}

# Called DEFINEDNAME by OpenOffice.
sub decodeLbl($$)
{
  my ($bytes, $version) = @_;

  CORE::state $builtinNames = 
  [
    'Consolidate_Area',
    'Auto_Open',
    'Auto_Close',
    'Extract',
    'Database',
    'Criteria',
    'Print_Area',
    'Print_Titles',
    'Recorder',
    'Data_Form',
    'Auto_Activate',
    'Auto_Deactivate',
    'Sheet_Title',
    '_FilterDatabase',
  ];

  my $decoder = StructDecoder->new($bytes);
  $decoder->registerType('string', ($version <= Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF5) ? \&decodeNoLenAnsiString : \&decodeNoLenXLUnicodeString);

  my $lbl = $decoder->decodeHash(['flags:u16', 'chKey:u8', 'nameLength:u8', 'formulaSize:u16', ':bytes[2]', 'sheetIndex:u16', 'menuLength:u8', 'descLength:u8', 'helpLength:u8', 'statusLength:u8', 'name:string[nameLength]', 'formula:bytes[formulaSize]']);

  # NOTE: We skip these fields entirely if the length is zero, because
  #  decodeNoLenXLUnicodeString would still read a flags byte otherwise.

  if ($lbl->{menuLength} > 0)
  {
    $lbl->{menu} = $decoder->decodeField('string', $lbl->{menuLength});
  }

  if ($lbl->{descLength} > 0)
  {
    $lbl->{desc} = $decoder->decodeField('string', $lbl->{descLength});
  }

  if ($lbl->{helpLength} > 0)
  {
    $lbl->{desc} = $decoder->decodeField('string', $lbl->{helpLength});
  }

  if ($lbl->{statusLength} > 0)
  {
    $lbl->{desc} = $decoder->decodeField('string', $lbl->{statusLength});
  }

  if (($lbl->{flags} & 0x20) && ($lbl->{nameLength} == 1))
  {
    $lbl->{name} = $builtinNames->[ord($lbl->{name})];
  }

#  my $lbl;
#  if ($version <= Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF5)
#  {
#    $lbl = $decoder->decodeHash(['flags:u16', 'chKey:u8', 'nameLength:u8', 'formulaSize:u16', ':bytes[2]', 'sheetIndex:u16', 'menuLength:u8', 'descLength:u8', 'helpLength:u8', 'statusLength:u8']);
#    $lbl->{name} = $decoder->decodeNoLenXLUnicodeString($lbl->{nameLength});
#  }
#  else
#  {
#    $lbl = $decoder->decodeHash(['flags:u16', 'chKey:u8', 'nameLength:u8', 'formulaSize:u16', ':bytes[2]', 'sheetIndex:u16', ':bytes[4]']);
#    $lbl->{name} = $decoder->decodeNoLenXLUnicodeString($lbl->{nameLength});
#  }

  return $lbl;
}

sub decodeIndex($$)
{
  my ($bytes, $version) = @_;

  my $decoder = StructDecoder->new($bytes);

  my $index;
  if ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF5)
  {
    $index = $decoder->decodeHash([':u32', 'minRow:u16', 'maxRow:u16', 'colWidthOffset:u32']);
  }
  elsif ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF8)
  {
    $index = $decoder->decodeHash([':u32', 'minRow:u32', 'maxRow:u32', 'colWidthOffset:u32']);
  }
  else
  {
    ...;
  }

  # NOTE: In a blank sheet, we can have minRow == maxRow == 0
  my $rowCount = $index->{maxRow} - $index->{minRow};  # NOTE: maxRow is actually the first unused trailing row
  #my $blockCount = int(($rowCount - 1) / 32) + 1;  # 32 rows per row block
  my $blockCount = ($rowCount + (Spreadsheet::Nifty::XLS::ROW_BLOCK_SIZE - 1)) >> Spreadsheet::Nifty::XLS::ROW_BLOCK_SHIFT;
  ($decoder->bytesLeft() >= ($blockCount * 4)) || die("decodeIndex: Expected array of at least ${blockCount} elements");

  #$index->{dbcellOffsets} = $decoder->decodeArray('u32', $decoder->bytesLeft() >> 2);
  $index->{dbcellOffsets} = $decoder->decodeArray('u32', $blockCount);

  return $index;
}

sub decodeDBCell($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);

  my $dbcell = {};
  $dbcell->{firstRowOffset} = $decoder->decodeField('u32');
  my $count = $decoder->bytesLeft() >> 1;
  ($count <= 32) || die("decodeDBCell: Row block larger than 32 rows?");
  $dbcell->{firstCellOffsets} = $decoder->decodeArray('u16', $count);
  return $dbcell;
}

sub decodeDefaultRowHeight($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $defaultRowHeight = $decoder->decodeHash(['flags:u16']);

  if ($decoder->bytesLeft())
  {
    $defaultRowHeight->{emptyRowHeight} = $decoder->decodeField('u16');
  }

  if ($decoder->bytesLeft())
  {
    $defaultRowHeight->{hiddenRowHeight} = $decoder->decodeField('u16');
  }

  return $defaultRowHeight;
}

sub decodeWsBool($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $wsBool = $decoder->decodeHash(['flags:u16']);
  return $wsBool;
}

sub decodeColInfo($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $colInfo = $decoder->decodeHash(['start:u16', 'end:u16', 'width:u16', 'ixfe:u16', 'flags:u16']);
  return $colInfo;
}

sub decodeAutoFilter($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $autoFilter = $decoder->decodeHash(['entry:u16', 'flags:u16']);
 
  sub decodeOperator($)
  {
    my ($decoder) = @_;

    my $operator = $decoder->decodeHash(['valueType:u8', 'comparison:u8']);

    # NOTE: All values seem to be padded to at least 8 bytes
    if ($operator->{valueType} == 0x00)  # Undefined
    {
      $decoder->getBytes(8);
    }
    elsif ($operator->{valueType} == 0x02)  # Rk
    {
      $operator->{value} = translateRk($decoder->decodeField('u32'));
      $decoder->getBytes(4);
    }
    elsif ($operator->{valueType} == 0x04)  # Xnum
    {
      $operator->{value} = $decoder->decodeField('f64');
    }
    elsif ($operator->{valueType} == 0x06)  # String
    {
      $decoder->getBytes(4);
      $operator->{strLen}   = $decoder->decodeField('u8');
      $operator->{strFlags} = $decoder->decodeField('u8');
      $decoder->getBytes(2);
    }
    elsif ($operator->{valueType} == 0x08)   # Bool/Error
    {
      $operator->{value}        = $decoder->decodeField('u8');
      $operator->{valueSubtype} = $decoder->decodeField('u8');
      $decoder->getBytes(6);
    }
    elsif ($operator->{valueType} == 0x0C)  # Blanks
    {
      $decoder->getBytes(8);
    }
    elsif ($operator->{valueType} == 0x0D)  # Non-blanks
    {
      $decoder->getBytes(8);
    }
 
    return $operator;
  }

  $autoFilter->{op1} = decodeOperator($decoder);
  $autoFilter->{op2} = decodeOperator($decoder);

  if ($autoFilter->{op1}->{valueType} == 0x06)
  {
    #printf("Reading op1 string size %d\n", $autoFilter->{op1}->{strLen});
    $autoFilter->{op1}->{value} = decodeNoLenXLUnicodeString($decoder, $autoFilter->{op1}->{strLen});
  }

  if ($autoFilter->{op2}->{valueType} == 0x06)
  { 
    #printf("Reading op2 string size %d\n", $autoFilter->{op2}->{strLen});
    $autoFilter->{op2}->{value} = decodeNoLenXLUnicodeString($decoder, $autoFilter->{op2}->{strLen});
  }

  return $autoFilter;
}

sub decodeDimensions($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $size = $decoder->bytesLeft();

  my $dimensions;
  if ($size == 10)  # BIFF7 and earlier?
  {
   $dimensions = $decoder->decodeHash(['minRow:u16', 'maxRow:u16', 'minCol:u16', 'maxCol:u16']);
  }
  elsif ($size == 14)  # BIFF8
  {
   $dimensions = $decoder->decodeHash(['minRow:u32', 'maxRow:u32', 'minCol:u16', 'maxCol:u16']);
  }
  else
  {
    die("Unexpected Dimensions record size of $size");
  }

  return $dimensions;
}

sub decodeRow($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $row = $decoder->decodeHash(['row:u16', 'minCol:u16', 'maxCol:u16', 'rowHeight:u16', ':u32', 'flags1:u16', 'flags2:u16']);
  return $row;
}

sub decodeBlank($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $cell = $decoder->decodeHash(['row:u16', 'col:u16', 'xf:u16']);
  return $cell;
}

sub decodeLabel($$)
{
  my ($bytes, $version) = @_;

  my $decoder = StructDecoder->new($bytes);
  $decoder->registerType('string', ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF8) ? \&decodeXLUnicodeString : \&decodeLen16AnsiString);

  my $cell = $decoder->decodeHash(['row:u16', 'col:u16', 'xf:u16', 'str:string']);
  return $cell;
}

sub decodeRString($$)
{
  my ($bytes, $version) = @_;

  my $decoder = StructDecoder->new($bytes);
  $decoder->registerType('string', ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF8) ? \&decodeXLUnicodeString : \&decodeLen16AnsiString);

  my $cell = $decoder->decodeHash(['row:u16', 'col:u16', 'xf:u16', 'str:string']);
  # TODO: The remaining bytes contain "rich text" runs
  return $cell;
}

sub decodeLabelSst($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $cell = $decoder->decodeHash(['row:u16', 'col:u16', 'xf:u16', 'si:u32']);
  return $cell;
}

sub decodeRK($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $cell = $decoder->decodeHash(['row:u16', 'col:u16', 'xf:u16', 'rk:u32']);
  return $cell;
}

sub decodeNumber($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $cell = $decoder->decodeHash(['row:u16', 'col:u16', 'xf:u16', 'num:f64']);
  return $cell;
}

sub decodeBoolErr($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $cell = $decoder->decodeHash(['row:u16', 'col:u16', 'xf:u16', 'value:u8', 'datatype:u8']);
  return $cell;
}

sub decodeMulBlank($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $mulBlank = $decoder->decodeHash(['row:u16', 'minCol:u16']);
  my $count = ($decoder->bytesLeft() - 2) >> 1;
  $mulBlank->{xfs} = $decoder->decodeArray('u16', $count);
  $mulBlank->{maxCol} = $decoder->decodeField('u16');

  # Sanity check
  (($mulBlank->{maxCol} - $mulBlank->{minCol} + 1) == $count) || die("Strange MulBlank array count, expected $count, got " . ($mulBlank->{maxCol} - $mulBlank->{minCol} + 1));
  
  return $mulBlank;
}

sub decodeMulRk($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $mulRk = $decoder->decodeHash(['row:u16', 'minCol:u16']);

  # Then follows an array of RkRecs, each 6 bytes, followed by maxCol:u16
  my $count = ($decoder->bytesLeft() - 2) / 6;
  my $recs = [];
  for (my $i = 0; $i < $count; $i++)
  { 
    my $rkRec = $decoder->decodeHash(['ixfe:u16', 'rk:u32']);
    push(@{$recs}, $rkRec);
  }

  $mulRk->{recs} = $recs;

  # Final field
  $mulRk->{maxCol} = $decoder->decodeField('u16');

  # Sanity check
  #printf("MulRk: Mincol %d maxcol %d count %d\n", $mulRk->{minCol}, $mulRk->{maxCol}, $count);
  (($mulRk->{maxCol} - $mulRk->{minCol} + 1) == $count) || die("Strange MulRk array count, expected $count, got " . ($mulRk->{maxCol} - $mulRk->{minCol} + 1));

  return $mulRk;
}

sub decodeFormula($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $cell = $decoder->decodeHash(['row:u16', 'col:u16', 'xf:u16', 'result:bytes[8]', 'flags:u16', ':u32']);
  my $count = $decoder->decodeField('u16');
  $cell->{formula} = $decoder->getBytes($count);
  $cell->{extra}   = $decoder->getBytes($decoder->bytesLeft());
  return $cell;
}

sub decodeShrFmla($)
{
  my ($bytes) = @_;

  my $decoder = StructDecoder->new($bytes);
  my $shrFmla = $decoder->decodeHash(['minRow:u16', 'maxRow:u16', 'minCol:u8', 'maxCol:u8', ':u8', 'useCount:u8', 'formulaSize:u16', 'formula:bytes[formulaSize]']);
  $shrFmla->{extra} = $decoder->getBytes($decoder->bytesLeft());

  return $shrFmla;
}

# A String record follows a Formula record when the result of the formula is a string.
# There is only a single field, so we just return the string value directly.
sub decodeString($$)
{
  my ($bytes, $version) = @_;

  my $decoder = StructDecoder->new($bytes);
  #  #$decoder->registerType('XLUnicodeString', \&decodeXLUnicodeString);
  #  $decoder->registerType('string', ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF8) ? \&decodeXLUnicodeString : \&decodeLen16AnsiString);
  #  my $string = $decoder->decodeField('string');
  my $string = (($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF8) ? \&decodeXLUnicodeString : \&decodeLen16AnsiString)->($decoder);
  return $string;
}

1;
