#!/usr/bin/perl -w
use warnings;
use strict;

# This is for "Office Binary Document RC4 Encryption", not to be confused with
#  "Office Binary Document RC4 CryptoAPI Encryption". In other words, this is
#  the non-CryptoAPI version.
package Spreadsheet::Nifty::XLS::Crypto::RC4;

use Digest::MD5;
use Crypt::RC4;

# The influence of the password is truncated to 40 bits. Could be for export
#  control reasons.

# Stream is re-keyed on every block boundary.
my $blockSize = 1024;

# === Class methods ===

sub new($$)
{
  my $class = shift();
  my ($password, $salt) = @_;

  (length($salt) == 16) || die("Salt must be 128 bits");

  my $self = {};
  $self->{password} = $password;
  $self->{salt}     = $salt;
  $self->{prepared} = undef;  # 40-bit key prepared from the password
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
  $new->{password} = $self->{password};
  $new->{salt}     = $self->{salt};
  $new->{prepared} = $self->{prepared};
  bless($new, ref($self));

  return $new;
}

# Using the password and salt, prepares a 40-bit hash value. These prepertory
#  steps do not use the block number, so they can be calculated once for a
#  given password/salt pair.
sub prepare()
{
  my $self = shift();

  my $passbytes = Encode::encode('UTF-16LE', $self->{password});
  #printf("Password UTF16 bytes: %s\n", unpack('H*', $passbytes));

  my $h = Digest::MD5::md5($passbytes);
  #printf("MD5 of password bytes: %s\n", unpack('H*', $h));

  my $intermediate = substr($h, 0, 5) . $self->{salt};
  (length($intermediate) == 21) || die();

  $intermediate x= 16;
  #printf("Intermediate: %s\n", unpack('H*', $intermediate));

  $h = Digest::MD5::md5($intermediate);
  #printf("MD5 of intermediate: %s\n", unpack('H*', $h));

  $self->{prepared} = substr($h, 0, 5);
  return;
}

# Given a block number, returns the key to be used for that block
sub keyForBlock($)
{
  my $self = shift();
  my ($block) = @_;

  my $input = $self->{prepared} . pack('V', $block);
  #printf("Truncate and concat with block: %s\n", unpack('H*', $input));

  my $key = Digest::MD5::md5($input);
  #printf("Final key for block %d: %s\n", $block, unpack('H*', $key));

  return $key;
}

# Returns true if the password matches the verifier values.
sub checkPassword($$)
{
  my $self = shift();
  my ($encryptedVerifier, $encryptedVerifierHash) = @_;

  #printf("encrypted verifier: %s\n", unpack('H*', $encryptedVerifier));
  #printf("encrypted verifierHash: %s\n", unpack('H*', $encryptedVerifierHash));

  my $key = $self->keyForBlock(0);
  my $rc4 = Crypt::RC4->new($key);

  # Decrypt the verifier and verifierHash
  my $verifier = $rc4->RC4($encryptedVerifier);
  my $verifierHash = $rc4->RC4($encryptedVerifierHash);
  
  #printf("verifier: %s\n", unpack('H*', $verifier));
  #printf("verifierHash: %s\n", unpack('H*', $verifierHash));

  # The hash of the decrypted verifier should match the decrypted verifierHash.
  my $check = Digest::MD5::md5($verifier);
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
