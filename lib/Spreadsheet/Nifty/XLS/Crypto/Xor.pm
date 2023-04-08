#!/usr/bin/perl -w
use warnings;
use strict;

# This is for "XOR Obfuscation".
package Spreadsheet::Nifty::XLS::Crypto::Xor;

# Initializers for preparing "xor key", based on password length
my $initialCodes = [ 0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3 ];

# Permutations for each round of calculating the "xor key".
# Passwords can be up to 15 characters in length, and there are seven rounds
#  per character (one for each ASCII character bit), giving us 105 values.
my $xorMatrix =
[
 0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09,
 0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF,
 0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0,
 0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40,
 0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5,
 0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A,
 0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9,
 0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0,
 0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC,
 0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10,
 0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168,
 0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C,
 0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD,
 0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC,
 0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4
];

# Used for calculating the xor buffer.
my $padArray = [ 0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00 ];

my $blockSize = 16;

# === Class methods ===

sub new($$)
{
  my $class = shift();
  my ($password) = @_;

  (length($password) > 15) && die("Overlong password");

  my $self = {};
  $self->{password} = $password;
  $self->{xorSeq} = undef;  # 128-bit xor sequence
  bless($self, $class);

  $self->prepare();

  return $self;
}

# Given a password, returns a 16-bit verifier hash value.
sub hashPasswordVerifier($)
{
  my $class = shift();
  my ($password) = @_;

  $password = Encode::encode('ASCII', $password);
  my $passwordLength = length($password);
  ($passwordLength > 15) && die("Overlong password");

  my $hash = 0;
  for (my $i = 0; $i < $passwordLength; $i++)
  {
    my $c = substr($password, $i, 1);
    my $o = ord($c);

    # Rotate left through low 15 bits, rotating one bit per character position
    my $v = $o;
    for (my $j = 0; $j <= $i; $j++)
    {
      $v = (($v << 1) & 0x7FFF) + ($v >> 14);
    }

    $hash ^= $v;
  }

  $hash ^= length($password);
  $hash ^= 0xCE4B;

  return $hash;
}

# Given a password, returns a 16-bit xor key.
sub calculateXorKey($)
{
  my $class = shift();
  my ($password) = @_;

  $password = Encode::encode('ASCII', $password);
  my $passwordLength = length($password);
  ($passwordLength > 15) && die("Overlong password");

  my $xorKey = $initialCodes->[$passwordLength - 1];
  my $matrixIndex = scalar(@{$xorMatrix}) - 1;

  # Loop through each character
  for (my $i = $passwordLength - 1; $i >= 0; $i--)
  {
    my $c = ord(substr($password, $i, 1));

    # Loop through the 7 bits of this ASCII character, and xor in the value from the matrix if set
    for (my $b = 6; $b >= 0; $b--)
    {
      #printf("Char %d bit %d value %d matrixIndex %d\n", $i, $b, $c, $matrixIndex);
      if (($c >> $b) & 0x01)
      {
        $xorKey ^= $xorMatrix->[$matrixIndex];
      }

      $matrixIndex--;
    }
  }

  return $xorKey;
}

# Given a password, returns a 128-bit (16-byte) xor sequence.
sub calculateXorSequence($)
{
  my $class = shift();
  my ($password) = @_;

  $password = Encode::encode('ASCII', $password);
  my $passwordLength = length($password);
  ($passwordLength > 15) && die("Overlong password");

  my $xorKey = $class->calculateXorKey($password);
  my $xorKeyLow  = $xorKey & 0xFF;
  my $xorKeyHigh = $xorKey >> 8;

  my $buf = [ map({ ord($_) } split(//, $password)), @{$padArray}[0..(16 - $passwordLength)] ];
  #print main::Dumper('START BUF', $buf);

  for (my $i = 0; $i < 16; $i += 2)
  {
    $buf->[$i]     ^= $xorKeyLow;
    $buf->[$i + 1] ^= $xorKeyHigh;
  }

  # Rotate each byte left by 2 bits
  for (my $i = 0; $i < 16; $i++)
  {
    my $v = $buf->[$i];
    $buf->[$i] = (($v << 2) & 0xFF) | ($v >> 6);
  }

  #print main::Dumper('END BUF', $buf);
  return $buf;
}

# === Instance methods ===

# Duplicates this decryptor object.
sub dup()
{
  my $self = shift();

  my $new = {};
  $new->{password} = $self->{password};
  $new->{xorSeq}   = $self->{xorSeq};
  bless($new, ref($self));

  return $new;
}

sub prepare()
{
  my $self = shift();

  $self->{xorSeq} = $self->calculateXorSequence($self->{password});
  return;
}

# Returns true if the password matches the verifier value.
sub checkPassword($$)
{
  my $self = shift();
  my ($verifier) = @_;

  my $check = $self->hashPasswordVerifier($self->{password});

  printf("xor verifier: %X\n", $verifier);
  printf("check       : %X\n", $check);

  return ($check eq $verifier);
}

sub decryptBiffPayload($$$$)
{
  my $self = shift();
  my ($offset, $type, $length, $encryptedPayload) = @_;

  my $payloadLength = length($encryptedPayload);
  ($length == $payloadLength) || die("Encrypted payload length does not match BIFF record length");

  #printf("Offset %d type %d (0x%04X) length %s\n", $offset, $type, $type, $length);
  #printf("  encryptedPayload: %s\n", unpack('H*', $encryptedPayload));

  my $seqIndex = $offset + $length;

  my $payload = '';
  for (my $i = 0; $i < $length; $i++)
  {
    my $b = ord(substr($encryptedPayload, $i, 1));
    #printf("    byte %d (0x%X)", $b, $b);
    $b = (($b << 3) & 0xFF) | ($b >> 5);
    #printf("    rotated %d (0x%X)\n", $b, $b);
    $payload .= chr($b ^ $self->{xorSeq}->[$seqIndex % $blockSize]);
    $seqIndex++;
  }

  if ($type == 0x0085)
  {
    # Special handling for BoundSheet8 - The first 4 bytes are never encrypted
    $payload = substr($encryptedPayload, 0, 4) . substr($payload, 4);
  }

  #printf("  payload         : %s\n", unpack('H*', $payload));

  return $payload;
}

1;
