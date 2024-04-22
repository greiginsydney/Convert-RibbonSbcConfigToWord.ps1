# Convert-RibbonSbcConfigToWord.ps1

The name's a bit of a mouthful, but "Convert-RibbonSbcConfigToWord.ps1" takes the backup file from your Sonus/Ribbon SBC 1000/2000/SWeLite gateway and creates a new Word document, with all of the important(?) configuration information captured  in tables.&nbsp;
It started life as a way to save the tedium of screen-scraping lots of fixed frames for my as-built documents, but it quickly became apparent that it would also make a useful tool for the offline review of a gateway's config (although it ain't  no "UxBuilder").

&nbsp;

<h2><span style="color: #ff0000;">Convert your backups from this:</h2>

<img src="https://user-images.githubusercontent.com/11004787/81054574-64b12f80-8f0a-11ea-8a57-60fef112945a.png" alt="" width="600" />

<h2><span style="color: #ff0000;">&hellip; to this:</h2>

<img src="https://user-images.githubusercontent.com/11004787/81054622-7abef000-8f0a-11ea-9148-28a696bd805d.png" alt="" width="600" />

## Features

- Decodes SweLite, 1k & 2k gateway versions up to 12.0.1. 

- Converts and saves the config to a new Word document, and optionally makes a PDF of it too. 

- Doesn't need any Lync PowerShell modules - it's fine on your Win7-Win10 machine (which needs to have Word installed). 

- Tested on Word 2010 through 2021 inclusive.

- Reasonably well debugged throughout (although time will tell of course!). 

- It's easy enough to turn off an unwanted section if you're only interested in a specific subset of the functionality, or are adapting to suit your environment. 

- Use the "RedactIP" switch to strip all IP addresses from the resulting file. 

- A "-SkipWrite" switch skips the relatively slow process of writing to Word. I added this for debugging and have left it in the released version to help anyone refining or enhancing the script. 

- Digicert kindly provide me with a code-signing certificate, so you don't need to drop your ExecutionPolicy to "Unrestricted". This also means you can be confident the code's not changed since it left here. 

- Uses your existing Word document template for the formatting of the table headings and table of contents. 

## Limitations

- It's not the fastest to run! Stuffing values into Word is a relatively slow process. With all options enabled allow up to 10 minutes for your average gateway. I've stolen Scott Hanselman's "<a href="http://www.hanselman.com/blog/ProgressBarsInPowerShell.aspx" target="_blank">Progress  Bars in PowerShell</a>" to help the watched pot boil. 

- If you find I've not captured a particularly  valuable element of the config, feel free to revise it to suit your needs - and send me the new bits to incorporate, please. If you're not feeling up to it, send me a sample config file and some screen-grabs of the bit you'd like to see added. 

- There are MANY hash tables throughout. (This is where an integer is stored in the config file, but it's decoded to text on-screen). I've put a lot of effort into capturing these, but there are going to be examples where I've missed one.  Please send me screen-grabs and sample config files and I'll update the script. 

- I'm largely letting Word control the width of the table columns, and you'll find many instances where it's done a poor job. You may need to edit every document the script produces to resolve this. 

## Instructions

* Backup your gateway, saving the config file somewhere accessible. Older gateway versions (pre 2.2.1) will save to an archive called "backup.tar.gz", while newer versions save as "SBC_Config_<hostname>_<version>_<date>.tar" 

* Older backups you'll need to unzip <em>twice</em> to recover the "symphonyconfig.xml" document we need, whilst newer ones only need one unzip to reveal their goodness <br /> >> If you have <a href="http://www.7-zip.org/" target="_blank">7-zip</a> installed you no longer need to unzip "SBC_Config_<hostname>_<version>_<date>.tar" - just feed the .tar file as the "-inputfile" 

* Open a PowerShell window 

* Now feed the backup into the script! Either place the config file in the same path as the script, give the full path in the command line, or a relative path:

```powershell 
PS H:\> Convert-RibbonSbcConfigToWord.ps1 -InputFile "symphonyconfig.xml" -OutputFile "CustomerSBC-config.docx"
```

If you don't specify an InputFile, the script goes looking for "symphonyconfig.xml" in its directory. If you don't specify an OutputFile, it will re-use the name of the InputFile and save as ".docx", overwriting any existing file of  the same name in that directory. If you add the "-MakePdf" switch, it will save the same output filename as a .pdf as well 

You can easily batch it:

```powershell
PS H:\> Get-ChildItem "d:\path\*.xml" -recurse | .\Convert-RibbonSbcConfigToWord.ps1 -IncludeNodeInfo -MakePdf -SkipUpdateCheck
```

```powershell
PS H:\> Get-ChildItem "d:\path\*.tar" -recurse | .\Convert-RibbonSbcConfigToWord.ps1 -MakePdf -SkipUpdateCheck
```

## Revision History

### (Read about older versions <a href="https://greiginsydney.com/uxbuilder" target="_blank"> on my blog</a>)

#### v12.1.0 22nd April 2024
- Added new bits in 12.1.0:
    - New TLS ciphers added to Security / TLS Profiles / Client & Server Cipher Lists
    - 'OPTIONS mode' added to SIP SigGps
    - 'TcpKeepAlive' to SIP Server Tables
    - 'OptionPassthrugh[SIC]' to SIP Profiles
    - New protocols to $TlsProtocolLookup (for TLS Profiles)
    - 'Server Cipher List' in TLS profiles
- Fixed bugs:
    - '<n/a> this rls' :-)

####  v12.0.1 3rd October 2023
- Added new bits found recently:
    - New values to $NumberTypeLookup, and renamed to $ClgNumberTypeLookup
    - Cloned $NumberTypeLookup to new $CldNumberTypeLookup, as some of the values no longer align between calling and called.
    - New values to $TransferCapabilityLookup, $NumberingPlanLookup
    - 'Normal' call type renamed 'Standard' in $CRDestinationTypeLookup & $CRCallPriorityLookup
    - Added 'Trunk Group' as a Destination Type in call routing table entry, and suppressed display of SigGp and other parameters
- Fixed bugs:
    - A Routing Table's top 'H-table' may incorrectly show a value for First Signaling Group when DestinationType is 'Deny' or 'Trunk Group'

#### 		v12.0.0 * Not released.
- Added new bits in 12.0.0:
    - 'Global Date/Time Format' in Logging Configuration / Global Log
    - 'Session Timer Offset' in SIP Profiles. (Appears if Session Timer is 'Enable')
    - Added 'Diagnostic Logging / By Call Criteria'
- Added new bits found recently:
    - New transformation types added to $InputFieldLookup & $OutputFieldLookup (thanks Mike)
    - Updated Logging Configuration / Global Log to show for all machine types
    - Added new 1k 'mirrored port' values to $PortMirrorPortLookup

#### v11.0.0 4th February 2022
- Added new bits in 11.0.0:
  - 'Enforce SG Codec Priority' in Media List
  - New section 'Listen Port' in SIP
  - Changed how SIP Sig Gps and SIP Recorders display 'Listen Port' info. (If new format is present it will be used, else legacy)
- Added new bits in 9.0.7:
  - Added new RSA-AES-GCM cipher suites for TLS 1.2 interop
  - Added new 'Media Codec Latch' in SIP Sig SPs
- Added [System.Version] declaration in Get-UpdateInfo to prevent issues where '9.0.4' is apparently > '11.0.0'
- Fixed bugs:
  - Routing table 'call priority' = 'Emergency' was being reported as 'blank' due to incorrect ID in $CRCallPriorityLookup

#### v9.0.7 * Not released.
- v11 preceeded 9.0.7 due to the staggered versioning of the soft and hard platforms.

#### v9.0.6 * Not released.
- No new functionality, no bugs unearthed.

#### v9.0.5 * Not released.
- No new functionality, no bugs unearthed.

#### v9.0.4B * Not released.
- Updated for PowerShell v7 compatibility:
  - Replaced all tests for "if -eq ''" with 'if ([string]::IsNullOrEmpty...'
  - Added '[char[]]' to multiple-value '.split()' methods
  - Changed $NodeInfoArray creation from '.split()' to '-split' & added blank line test/continue
  - Removed reference to [Microsoft.Office.Interop.Word.WdExportCreateBookmarks] enum
- Removed obsolete $NotificationData

#### v9.0.4 15th July 2021
- Added new bit in 9.0.4:
  - Added Global Security Options / Test a Call

#### v9.0.3 * Not released.
- No new functionality, no bugs unearthed.

#### v9.0.2 28th January 2021
- Changed Network Monitoring / Link monitors from h-table to v-table and updated to accommodate new values in 9.0.2
- Added Protocols / IPSec / Tunnel Table
- Changed some references to '<n/a>' and '<Not Captured>' to '<Not Available>' to reduce ambiguity & increase consistency
- Now stamps '<Not Available>' into System / Node-Level Settings if the SweLite ID in nodeinfo.txt is blank
- Fixed bugs:
  - Fixed bug introduced in 9.0.1 where the lower half of each SIP Server table was skipped. (Bad test of SIP Recorders)
  - Suppressed display of 'Send STUN Packets' in Media / Media System Configuration for the SweLite

<br>&nbsp;

### Help me improve the script

PLEASE let me know if you encounter any problems. All I ask is a copy of the .tar or symphonyconfig.xml file (de-identified as you wish), and a screen-grab from the browser to show me what it's <em>meant</em> to look like on-screen.

<hr />

<br>

\- G.

<br>

This script was originally published at [https://greiginsydney.com/uxbuilder/](https://greiginsydney.com/uxbuilder/).
