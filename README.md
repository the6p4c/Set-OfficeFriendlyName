# `Set-OfficeFriendlyName`
Sets the friendly name used by an Office identity (that is, the name which appears in the top right corner of all Office applications).

![](screenshot.png)

## Usage
```
PS > Set-OfficeFriendlyName "Your Name"
Found 1 Office identity
No identity specified: updating only Office identity 1a859458-dc27-4be1-95fd-1dd644d21a8f_ADAL
Setting new friendly name to Your Name
Done! Note: this change will likely be undone by signing in again.
```
```
PS > Set-OfficeFriendlyName "Your Name" 1a859458-dc27-4be1-95fd-1dd644d21a8f_ADAL
Found 1 Office identity
Updating Office identity 1a859458-dc27-4be1-95fd-1dd644d21a8f_ADAL
Setting new friendly name to Your Name
Done! Note: this change will likely be undone by signing in again.
```
```
PS > Set-OfficeFriendlyName "New Name"
Found 2 Office identities
No identity specified and more than one Office identity was found.

Identity                                  FriendlyName
--------                                  ------------
1a859458-dc27-4be1-95fd-1dd644d21a8f_ADAL Your Name
dd6875ec-83d1-4429-a2c1-b49d5badb00f_ADAL Someone's Name

Exception: ...
Line |
  xx |      throw "You must specify an identity."
     |      ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
     | You must specify an identity.
```
