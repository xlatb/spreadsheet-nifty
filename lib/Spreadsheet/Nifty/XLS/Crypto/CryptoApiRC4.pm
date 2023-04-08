#!/usr/bin/perl -w
use warnings;
use strict;

# This is "Office Binary Document RC4 CryptoAPI Encryption". There is also a
#  non-CryptoAPI Version in a different package.
package Spreadsheet::Nifty::XLS::Crypto::CryptoApiRC4;

use Digest::SHA qw();
use Crypt::RC4 qw();

# Stream is re-keyed on every block boundary.
my $blockSize = 1024;

# === Class methods ===

sub new($$)
{
  my $class = shift();
  my ($password, $header, $verifier) = @_;

  #(length($salt) == 16) || die("Salt must be 128 bits");
  ($header->{encryptionType} == 0x6801) || die("Expected encryption type 0x6801 (RC4)");
  ($header->{hashType} == 0x8004) || die("Expected hash type 0x8004 (SHA-1)");
  ($header->{flags} & 0x4) || die("Expected bit 2 set in flags");  # CryptoAPI flag
  (($header->{flags} & 0x20) == 0) || die("Expected bit 6 unset in flags");  # AES flag
  (($header->{keySize} % 8) == 0) || die("Expected keySize to be a multile of 8");

  ($verifier->{saltSize} == 16) || die("Expected salt size 16");
  ($verifier->{verifierHashSize} == 20) || die("Expected hash size 20");  # SHA-1 hash size

  my $self = {};
  $self->{password} = $password;
  $self->{header}   = $header;
  $self->{verifier} = $verifier;
  $self->{prepared} = undef;
  bless($self, $class);

  $self->prepare();

  return $self;
}

# === Instance methods ===

# Duplicates this decryptor object.
sub dup()
{
  my $self = shift();

  my $new = {};
  for my $f (qw(password header verifier prepared))
  {
    $new->{$f} = $self->{$f};
  }

  bless($new, ref($self));
  return $new;
}

# Prepares an intermediate hash value from the password and salt. Does not use
#  a block number.
sub prepare()
{
  my $self = shift();

  my $passbytes = Encode::encode('UTF-16LE', $self->{password});
  #printf("Password UTF16 bytes: %s\n", unpack('H*', $passbytes));

  my $h = Digest::SHA::sha1($self->{verifier}->{salt} . $passbytes);
  #printf("SHA-1 of password and salt: %s\n", unpack('H*', $h));

  $self->{prepared} = $h;
  return;
}

# Given a block number, returns the key to be used for that block.
sub keyForBlock($)
{
  my $self = shift();
  my ($block) = @_;

  my $input = $self->{prepared} . pack('V', $block);
  #printf("Concat with block: %s\n", unpack('H*', $input));

  my $h = Digest::SHA::sha1($input);
  #printf("Final hash value for block %d: %s\n", $block, unpack('H*', $h));

  my $key;
  if (($self->{header}->{keySize} == 0) || ($self->{header}->{keySize} == 40))
  {
    # Special case for 40 bits. We pad the end with zeroes up to 128 bits.
    # NOTE: A keySize of zero is interpreted as 40.
    $key = substr($h, 0, 5) . ("\x00" x 11);
  }
  else
  {
    # The key is simply the first keySize bits of the hash
    $key = substr($h, 0, $self->{header}->{keySize} >> 3);
  }

  #printf("Final key for block %d: %s\n", $block, unpack('H*', $key));
  return $key;
}

# Returns true if the password matches the verifier values.
sub checkPassword($$)
{
  my $self = shift();

  #printf("encrypted verifier: %s\n", unpack('H*', $self->{verifier}->{encryptedVerifier}));
  #printf("encrypted verifierHash: %s\n", unpack('H*', $self->{verifier}->{encryptedVerifierHash}));

  my $key = $self->keyForBlock(0);
  my $rc4 = Crypt::RC4->new($key);

  # Decrypt the verifier and verifierHash
  my $verifier = $rc4->RC4($self->{verifier}->{encryptedVerifier});
  my $verifierHash = $rc4->RC4($self->{verifier}->{encryptedVerifierHash});
  
  #printf("verifier: %s\n", unpack('H*', $verifier));
  #printf("verifierHash: %s\n", unpack('H*', $verifierHash));

  # The hash of the decrypted verifier should match the decrypted verifierHash.
  my $check = Digest::SHA::sha1($verifier);
  #printf("check: %s\n", unpack('H*', $check));

  return ($check eq $verifierHash);
}

sub decryptBiffPayload($$$$)
{
  my $self = shift();
  my ($offset, $type, $length, $encryptedPayload) = @_;

  my $payloadLength = length($encryptedPayload);
  ($length == $payloadLength) || die("Encrypted payload length does not match BIFF record length");

  #printf("Offset %d type %d (0x%04X) length %s\n", $offset, $type, $type, $length);

  my $keystream = '';

  my $startBlock = int($offset / $blockSize);
  my $endBlock   = int(($offset + $length) / $blockSize);

  # TODO: Cache the rc4 object, block number, and offset. We could reuse it when the block
  #  number is the same and the offset increases.
  # TODO: Or even better, generate the keystream for the entire current block and cache that?
  # Get keystream for first block
  my $key = $self->keyForBlock($startBlock);
  my $rc4 = Crypt::RC4->new($key);
  my $runOffset = $offset % $blockSize;
  my $runMaxLen = $blockSize - $runOffset;
  my $runLen    = ($length > $runMaxLen) ? $runMaxLen : $length;
  #printf("runOffset %d runMaxLen %d runLen %d\n", $runOffset, $runMaxLen, $runLen);
  $rc4->RC4("\x00" x $runOffset);  # Discard the leading part of the keystream up to the current offset
  $keystream .= $rc4->RC4("\x00" x $runLen);
  $length -= $runLen;

  # Get keystream for remaining blocks
  for (my $b = $startBlock + 1; $length; $b++)
  {
    #printf("following block b %d length %d\n", $b, $length);
    $key = $self->keyForBlock($b);
    $rc4 = Crypt::RC4->new($key);
    $runLen = ($length > $blockSize) ? $blockSize : $length;
    $keystream .= $rc4->RC4("\x00" x $runLen);
    $length -= $runLen;
  }

  (length($keystream) == length($encryptedPayload)) || die("Unexpected keystream length");

  # XOR the encrypted payload with the keystream to get the unencrypted payload
  my $payload;
  if ($type == 0x0085)
  {
    # Special handling for BoundSheet8 - The first 4 bytes are never encrypted
    $payload = substr($encryptedPayload, 0, 4) . (substr($encryptedPayload, 4) ^ substr($keystream, 4));
  }
  else
  {
    $payload = $encryptedPayload ^ $keystream;
  }

  #printf("keystream %s\n", unpack('H*', $keystream));
  #printf("payload %s\n", unpack('H*', $payload));

  return $payload;
}

1;
