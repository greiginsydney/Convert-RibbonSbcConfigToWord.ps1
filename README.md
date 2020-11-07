# Convert-RibbonSbcConfigToWord.ps1

The name's a bit of a mouthful, but "Convert-RibbonSbcConfigToWord.ps1" takes the backup file from your Sonus/Ribbon SBC 1000/2000/SWeLite gateway and creates a new Word document, with all of the important(?) configuration information captured  in tables.&nbsp;
It started life as a way to save the tedium of screen-scraping lots of fixed frames for my as-built documents, but it quickly became apparent that it would also make a useful tool for the offline review of a gateway's config (although it ain't  no "UxBuilder").

&nbsp;

<h2><span style="color: #ff0000;">Convert your backups from this:</h2>

<img src="https://user-images.githubusercontent.com/11004787/81054574-64b12f80-8f0a-11ea-8a57-60fef112945a.png" alt="" width="600" />

<h2><span style="color: #ff0000;">&hellip; to this:</h2>

<img src="https://user-images.githubusercontent.com/11004787/81054622-7abef000-8f0a-11ea-9148-28a696bd805d.png" alt="" width="600" />

## Features

- Decodes SweLite, 1k & 2k gateway versions up to 9.0.1. 

- Converts and saves the config to a new Word document, and optionally makes a PDF of it too. 

- Doesn't need any Lync PowerShell modules - it's fine on your Win7-Win10 machine (which needs to have Word installed). 

- Tested on Word 2010, 2013 & 2016. 

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

#### v9.0.1 - 7th November 2020
- Consolidated creation of $SgTableLookup to new combined section in the 'first run' loop, as SigGps now recursively reference themselves for Call Recording
- Removed redundant Sig Gp type (e.g. '(SIP)') from the start of each SigGp heading
- Updated Security / TLS Profiles to now show only one instance of Certificate, now under 'Common Attributes'
- Added new bits in 9.0.0 (SweLite) / 9.0.0 (1k/2k):
	- 'SIP Recorder' (type) in SIP Server Tables
	- 'SIP Recording' section in SIP
	- 'SIP Recording' in SIP Sig Gps
	- 'E911 Notification Manager' in Emergency Services
	- Removed Forking from the licences table (as from v9 it's now licenced by default - but still displaying in the 1k/2k version of the table)
- Fixed bugs / new bits missed previously:
	- Added 'Drain' to $EnabledLookup & replaced 'Enabled' value with 'customAdminState' in Sig Gps where the latter is present
	- SIP SigGp was reporting the default $SipGp.IE.Description instead of $SIPgroupDescription. (Only apparent if $SipGp.IE.Description was blank)
	- Fixed bug where Security / TLS Profiles was showing a blank instead of the SBC certificate name

#### v8.1.5 16th August 2020
- Added display of SILK licences to the SweLite's System / Licensing / Current Licenses
- Now reports 'License Expiration' value from nodeinfo.txt for 1k/2k. (Was previously suppressed. I don't know why).
- Updated label on SweLite's License 'Virtual DSP Sessions' to 'Enhanced Media Sessions with Transcoding'
- Updated label on SweLite's License 'Proxy RTP <-> SRTP Sessions' to 'Enhanced Media Sessions without Transcoding'
- Removed test to suppress System / Node Level Settings / Country Level Information if SweLite: it's visible there now
- SweLite: Removed the 'License Expiration' line from the bottom of the licences table
- Changed license display: if not licenced, replace any '0' value with 'Not Licensed'
- Added new bits in 8.1.0 (SweLite) / 8.1.5 (1k/2k):
    - "Teams Local Media optimization" in SIP Sig Gps
    - Primary & Secondary Source to System / Node Level Settings / Domain Name Service
- Fixed bugs:
    - Updated input "$Fullname" to remove array declaration (Tks Mike.)
    - Some SBCs previously reported two blank lines between feature licences and the expiry, etc footer. Now back to just one.
    - SIP SigGp: Supported Audio Mode of "Proxy with local SRTP" & Proxy Local SRTP Crypto Profile ID should not show on a 1k/2k. Corrected.
					
#### v8.1.0B 18th March 2020

- Fixed bug where TT's were no longer arranged alphabetically - broken with the new layout introduced in v8.0.0. (Tks Mike.) 

#### v8.1.0 14th March 2020

- Restructured to support batch processing direct from the pipeline. Added new batch example:<br /> gci *.tar | .\Convert-RibbonSbcConfigToWord.ps1 -SkipUpdateCheck -MakePDF 

- Added new bits in 8.1.0: 

    - "Country Code" in System / Node Level Settings 

    - "Bad Actors" table under Settings / Security 

    - "Explicit Acknowledgement of Pre-Login Info" in Security / Users / Global Security Options 

    - "Certificate" in Security / TLS Profiles 

    - "SBC Supplementary Certificates" in Security / SBC Certificates 

    - "SILK" in System / Licensing / Current Licenses 

    - "DC Enabled" in Auth and Directory Services / Acive Directory / Domain Controllers 

- Fixed bugs: 

    - "Lockout duration" was always showing under Security / Users / Global Security Options, even if Number of Failed Logins was set to "No lockout" 

    - Was not correctly handling "SupportsShouldProcess". As resolving this will add a confirm prompt for each iteration (potentially breaking existing workflow), I've opted to disable it 

    - Under System / Node-Level Settings, &lsquo;Power LED' was misspelt 

- Updated label on SIP Sig Sp &lsquo;Outbound Proxy' to now show as &lsquo;Outbound Proxy IP/FQDN' 

#### v8.0.3 * Not released.

- Added display of settings when Security / Users / Global Security Options / Enhanced Password Security = Yes 

#### v8.0.2 * Not released.

- No new functionality, no bugs unearthed. 


### Help me improve the script

PLEASE let me know if you encounter any problems. All I ask is a copy of the .tar or symphonyconfig.xml file (de-identified as you wish), and a screen-grab from the browser to show me what it's <em>meant</em> to look like on-screen.

<hr />

<br>

\- G.

<br>

This script was originally published at [https://greiginsydney.com/uxbuilder/](https://greiginsydney.com/uxbuilder/).
