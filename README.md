# Add-PFReplica.ps1

Add legacy public folder replicas for multiple legacy public folder servers recursively

## Description

This script adds multiple public folder servers to all public folders below a TopPublicFolder.

The script has been developed as part of an on-premise legacy public folder migration from Exchange 2007 to Exchange Server 2010.

The script waits a given timespan in seconds to let public folder hierarchy replication and replica backfill requests kick in.

It is assumed that the script is being run in an Exchange 2007 or Exchange 2010 server.

## Updates

2015-09-24: Fixes to run properly with Exchange 2010

## Parameters

### ServersToAdd

String array containing the legacy public folders to add

### SecondsToPause

Seconds to pause between each step, this should be a reasonalbe timespan (default = 240)

### PublicFolderServer

The server name of the legacy public folder server to contact for changes

### TopPublicFolder

Name of the legacy top public folder

## Examples

``` PowerShell
.\Add-PFReplica.ps1 -ServersToAdd SERVER01,SERVER02 -PublicFolderServer SERVER01 -TopPublicFolder "\COMMUNICATIONS\PR
```

Add replicas for SERVER01,SERVER02 to all sub folders of \COMMUNICATIONS\PR

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)