function Set-OfficeFriendlyName {
<#
.SYNOPSIS
  Sets the friendly name used by an Office identity (that is, the name which appears in the top
  right corner of all Office applications).

.DESCRIPTION
  Set-OfficeFriendlyName sets the "friendly name" used by an Office identity. The friendly name is
  is the string which appears in the top right corner of all Office applications, and is stored
  per sign-in in the Windows registry.

  The initials-based profile icon is automatically generated and cannot be customised independently
  from the friendly name, but the initials will reflect the newly set friendly name (i.e.
  "Your Name" will generate a "YN" icon).

  Signing in again is likely to reset the friendly name to the account's default.

.PARAMETER NewFriendlyName
  A non-empty string to set as the new friendly name.

.PARAMETER Identity
  The unique ID associated with the identity to change. If an identity is not provided and only a
  single identity can be found, the friendly name of this identity will be changed. If more than one
  identity is found, and Identity is not specified, the script will list all discovered identities
  and exit with an error.

.EXAMPLE
  Set-OfficeFriendlyName "Your Name"

.EXAMPLE
  Set-OfficeFriendlyName "Your Name" 1a859458-dc27-4be1-95fd-1dd644d21a8f_ADAL

.NOTES
  Repository: https://github.com/the6p4c/Set-OfficeFriendlyName
#>
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory)]
    [ValidateLength(1, [int]::MaxValue)]
    [String] $NewFriendlyName,
    [String] $Identity
  )

  $ErrorActionPreference = "Stop"
  $IdentitiesPath = "HKCU:\Software\Microsoft\Office\16.0\Common\Identity\Identities"
  $Identities = @(Get-ChildItem -Path $IdentitiesPath)

  # Show the user the available Office identities if there isn't exactly one, and they didn't
  # specify which one to work with
  $NumIdentities = $Identities.Length
  if ($NumIdentities -eq 0) {
    throw "No Office identities found."
  }

  Write-Host "Found $NumIdentities Office $(if ($NumIdentities -eq 1) {"identity"} else {"identities"})"
  if (! $Identity -And $NumIdentities -ne 1) {
    $PrettyIdentities = $Identities | % { [PSCustomObject]@{
      Identity = Split-Path -Leaf $_.Name;
      FriendlyName = ($_ | Get-ItemProperty).FriendlyName | % { if ($_) { $_ } else { "<not set>" } }
    } }

    Write-Host -ForegroundColor Red "No identity specified and more than one Office identity was found."
    $PrettyIdentities | Format-Table
    throw "You must specify an identity."
  }

  # Tell the user which identity we're about to update
  $IdentityKey = if (! $Identity) {
    $IdentityKey = $Identities[0]
    Write-Host -ForegroundColor Yellow "No identity specified: updating only Office identity $(Split-Path -Leaf $IdentityKey)"
    $IdentityKey
  } else {
    $IdentityKeyPath = Join-Path -Path $IdentitiesPath -ChildPath $Identity
    if (! (Test-Path -PathType Container $IdentityKeyPath)) {
      throw "Could not find identity $Identity."
    }

    $IdentityKey = Get-Item $IdentityKeyPath
    Write-Host -ForegroundColor Yellow "Updating Office identity $Identity"
    $IdentityKey
  }

  # Update the identity
  Write-Host "Setting new friendly name to $NewFriendlyName"
  $IdentityKey | Set-ItemProperty -Name FriendlyName -Value $NewFriendlyName
  Write-Host -ForegroundColor Green "Done! " -NoNewline
  Write-Host -ForegroundColor Blue "Note: this change will likely be undone by signing in again."
}
