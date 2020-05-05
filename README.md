# Convert-RibbonSbcConfigToWord.ps1

The name's a bit of a mouthful, but "Convert-RibbonSbcConfigToWord.ps1" takes the backup file from your Sonus/Ribbon SBC 1000/2000/SWeLite gateway and creates a new Word document, with all of the important(?) configuration information captured  in tables.&nbsp;
It started life as a way to save the tedium of screen-scraping lots of fixed frames for my as-built documents, but it quickly became apparent that it would also make a useful tool for the offline review of a gateway's config (although it ain't  no "UxBuilder").

&nbsp;

<h2><span style="color: #ff0000;">Convert your backups from this:</h2>

<img src="https://user-images.githubusercontent.com/11004787/81054574-64b12f80-8f0a-11ea-8a57-60fef112945a.png" alt="" width="600" />

<h2><span style="color: #ff0000;">&hellip; to this:</h2>

<img src="https://user-images.githubusercontent.com/11004787/81054622-7abef000-8f0a-11ea-9148-28a696bd805d.png" alt="" width="600" />

## Features

- Decodes gateway versions up to 8.1.0 & SWe Lite. 

- Converts and saves the config to a new Word document, and optionally makes a PDF of it too. 

- Doesn't need any Lync PowerShell modules - it's fine on your Win7-Win10 machine (which needs to have Word installed). 

- Tested on Word 2010, 2013 & 2016. 

- Reasonably well debugged throughout (although time will tell of course!). 

- It's easy enough to turn off an unwanted section if you're only interested in a specific subset of the functionality, or are adapting to suit your environment. 

- Use the "RedactIP" switch to strip all IP addresses from the resulting file. 

- A "-SkipWrite" switch skips the relatively slow process of writing to Word. I added this for debugging and have left it in the released version to help anyone refining or enhancing the script. 

- From v2 the script has been signed, so no longer requires you to drop your ExecutionPolicy to "Unrestricted". This also means you can be confident the code's not changed since it left here. 

- Uses your existing Word document template for the formatting of the table headings and table of contents. 

## Limitations

- It's not the fastest to run! Stuffing values into Word is a relatively slow process. With all options enabled allow up to 10 minutes for your average gateway. I've stolen Scott Hanselman's "<a href="http://www.hanselman.com/blog/ProgressBarsInPowerShell.aspx" target="_blank">Progress  Bars in PowerShell</a>" to help the watched pot boil. 

- v8.0 captures even more of the gateway's config than before, however it still doesn't capture EVERYTHING. There's an enormous amount in there, and it's growing with each release! If you find I've not captured a particularly  valuable element of the config, feel free to revise it to suit your needs - and send me the new bits to incorporate, please. If you're not feeling up to it, send me a sample config file and some screen-grabs of the bit you'd like to see added. 

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

#### v8.0.1 - 30th January 2019

- Added new bits in 8.0.1: 

    - Changed ISDN Sig Gp for new 'Until Timer Expiry' option in 'Play Inband Message Post-Disconnect' 

    - Added new 'Convert DISC PI=1,8 to Progress' to ISDN Sig Gp 

- Fixed bugs: 

    - If the SBC has only 1 Transformation Table it would not display. (Bug in the alpha sort introduced in v7.0.0B) 

    - Corrected several bugs in display of 'DTLS-SRTP Profiles' 

    - Corrected bug in Media Lists where the 'DTLS-SRTP Profile' value wasn't correctly reported 

    - Certs was truncating parameters if one happened to be a URL (as I was splitting on '/') 

    - Certs wasn't populating the issuer's details (the RHS of the table) 

- Created new $SipProfileCntctHdrLookup & edited values in $SipProfileFrmHdrLookup for new on-screen text 

- Added "[No CN]" text if a cert (so far GoDaddy roots) has no CN 

#### v8.0.0B - 16th November 2018

- Updated label and display of SIP Sig Gp / Media Info / Supported Audio(/Fax) Modes to match current display syntax: Lite and 1k2k 

- Fixed bug: 

    - The logic for displaying SIP Sig Gp ICE Mode was faulty. (Mitsu-San yet again on the ball!) 

#### v8.0.0 - 26th September 2018

- Updated name to "Convert-RibbonSbcConfigToWord.ps1" 

- Added new bits in 8.0.0 (b502): 

    - NEW COLOUR SCHEME! (And added "-SonusGreen" switch for the die-hards) 

    - TOD Routing 

    - Load Balancing options 4 & 5 to $SipLoadBalancingLookup 

    - 'SIP Failover Cause Codes' in SIP Sig Gps 

    - 'RTCP Multiplexing' in SIP Sig Gps 

    - 'ICE Mode' in SIP Sig Gps 

    - Renamed 'Media Crypto Profiles' to 'SDES-SRTP Profiles' 

    - 'DTLS-SRTP Profiles' 

    - 'Redundancy Retry Timer' in SIP Profiles 

    - Grouped Transformations, TOD, Call Routing & Call Actions in new 'Call Routing' group to align with the new on-screen layout 

- Added new test in Test-ForNull to drop a marker in the doc if an expected lookup returned $null 

- Set Word to disable spelling and grammar checking in the document. (Thanks Mike!) 

- Hid the display of "Translations" (long obsolete) 

- Fixed bugs: 

    - Renamed variables in the SGPeers' .split() code to make them unique. (Had accidentally reused old having copy/pasted) 

    - Corrected display of 'SDES-SRTP Profiles' Derivation Rate to add the '2^' prefix for non-zero values 

    - If you ran with "-do calls" it would incorrectly not show the "incomplete" warning after the TOC. ("calls" sounds like "all") 

#### v7.0.3 - 31st July 2018

- Added new bits in 7.0.2 (b485) and SWe Lite 7.0.3 (b141): 

    - New SIP Sig Gp Interop Mode 'Header Transparency' 

    - Changed SIP Profile's 'FQDN in Contact Header' from $EnableLookup to $SipProfileFrmHdrLookup for Teams Direct Routing 

- Changed Media Profiles and Media Crypto Profiles to V-tables 

- Improved error reporting if the PDF exists and is open (i.e. locked) when the script goes to create it 

- Major update to the SIP Message Manipulation rules: 

    - Changed handling of SIP ElementDescriptors to break out the Token, Prefix and Suffix to a separate line for each 

    - Added the "pre-400" messages to $SIPDescriptionLookup 

    - Changed SMM rules to use $SIPDescriptionLookup when Applicable Messages = Selected Messages & a SIP code is matched 

    - SMM Header & Request Line rules were not displaying URI & URIUser parameters 

    - Changed SMM rules to display the header name in Title Case 

    - Added some missing descriptors to $SipElementDescElementClassLookup & $SipElementDescActionLookup to improve SMM 

    - Corrected display of "Ordinal" value in Header rules. Now only shows when the header name is Contact, Route, Record-Route or History-Info or PAI 

    - Suppressed incorrect display of "Value" when Action = Remove 

- Fixed bugs: 

    - Corrected SWeLite's Media Crypto Profiles, suppressing display of Master Key Lifetime, Lifetime Value and Derivation Rate values 

    - Func Dump-ErrToScreen was not correctly dumping the error to screen when in "non-debug" mode 

    - Added some missing SIP messages to the $SIPDescriptionLookup. (Thank you again Mitsu-San) 

    - Updated Strip-TrailingCR to keep removing the last char from the string until the last char ISN'T a CR. (Was only stripping once). 


#### v7.0.1 - 14th April 2018

- Added new bit in 7.0.1 (b483): 

    - "Encrypt AD Cache" in Auth and Directory Services / AD / Configuration 

- Re-jigged AD / Configuration to current layout. Removed some old variables. Changed label on "AD Backup" to "AD Backup Failure Alarm" 

### Help me improve the script

PLEASE let me know if you encounter any problems. All I ask is a copy of the .tar or symphonyconfig.xml file (de-identified as you wish), and a screen-grab from the browser to show me what it's <em>meant</em> to look like on-screen.

<hr />

<br>

\- G.

<br>

This script was originally published at [https://greiginsydney.com/uxbuilder/](https://greiginsydney.com/uxbuilder/).
