<#
.SYNOPSIS
	This script takes the .tar backup or XML config file from a Ribbon/Sonus/NET UX/SBC1000, 2000 or SWe Lite gateway's backup archive
	and saves a significant proportion of the content to a new Word document.

.DESCRIPTION
	This script takes the 'symphonyconfig.xml' file from a Ribbon/Sonus/NET UX/SBC1000, 2000 or SWe Lite gateway (extracted
	from the backup archive by the user) and decodes much of the XML therein, pasting relevant(?) configuration
	into tables in a new Word document.
	Word needs to be installed on this machine to run.
	It only uses standard PowerShell, calling no Lync/SfB commandlets.
	With 7-zip[.org] installed you can just feed the script the .tar file & it will extract 'symphonyconfig.xml' into memory

	It will run with no command-line parameters and assumes default values for the source and destination files.

.NOTES
	Version				: 9.0.2
	Date				: TBA January 2021
	Gateway versions	: 2.1.1 - 9.0.2
	Author				: Greig Sheridan
	There are lots of credits at the bottom of the script

	WISH-LIST / TODO:
					- OSPF
					- RIP
					- The reporting of licencing for the SWe Lite might need some more work. I'm not happy with how Transcoding is handled, or temp licence expiry.

	KNOWN ISSUES:
					- Logical Interfaces doesn't suppress display of Ethernet info for a paired port
					- 'Port ASM1' doesn't show under Node Interfaces / Ports unless NodeInfo.txt is provided, or extracting from .tar


	Revision History 	:
				v9.0.2 TBA January 2021
					Changed Network Monitoring / Link monitors from h-table to v-table to accommodate new values in 9.0.2
	
				v9.0.1 7th November 2020
					Consolidated creation of $SgTableLookup to new combined section in the 'first run' loop, as SigGps now recursively reference themselves for Call Recording
					Removed redundant Sig Gp type (e.g. '(SIP)') from the start of each SigGp heading
					Updated Security / TLS Profiles to now show only one instance of Certificate, now under 'Common Attributes'
					Added new bits in 9.0.0 (SweLite) / 9.0.0 (1k/2k):
						- 'SIP Recorder' (type) in SIP Server Tables
						- 'SIP Recording' section in SIP
						- 'SIP Recording' in SIP Sig Gps
						- 'E911 Notification Manager' in Emergency Services
						- Removed Forking from the licences table (as from v9 it's now licenced by default - but still displaying in the 1k/2k version of the table)
					Fixed bugs / new bits missed previously:
						- Added 'Drain' to $EnabledLookup & replaced 'Enabled' value with 'customAdminState' in Sig Gps where the latter is present
						- SIP SigGp was reporting the default $SipGp.IE.Description instead of $SIPgroupDescription. (Only apparent if $SipGp.IE.Description was blank)
						- Fixed bug where Security / TLS Profiles was showing a blank instead of the SBC certificate name

				v8.1.5 16th August 2020
					Added display of SILK licences to the SweLite's System / Licensing / Current Licenses
					Now reports 'License Expiration' value from nodeinfo.txt for 1k/2k. (Was previously suppressed. I don't know why).
					Updated label on SweLite's License 'Virtual DSP Sessions' to 'Enhanced Media Sessions with Transcoding'
					Updated label on SweLite's License 'Proxy RTP <-> SRTP Sessions' to 'Enhanced Media Sessions without Transcoding'
					Removed test to suppress System / Node Level Settings / Country Level Information if SweLite: it's visible there now
					SweLite: Removed the 'License Expiration' line from the bottom of the licences table
					Changed license display: if not licenced, replace any '0' value with 'Not Licensed'
					Added new bits in 8.1.0 (SweLite) / 8.1.5 (1k/2k):
						- "Teams Local Media optimization" in SIP Sig Gps
						- Primary & Secondary Source to System / Node Level Settings / Domain Name Service
					Fixed bugs:
						- Updated input "$Fullname" to remove array declaration (Tks Mike.)
						- Some SBCs previously reported two blank lines between feature licences and the expiry, etc footer. Now back to just one.
						- SIP SigGp: Supported Audio Mode of "Proxy with local SRTP" & Proxy Local SRTP Crypto Profile ID should not show on a 1k/2k. Corrected.

				v8.1.0B 18th March 2020
					Fixed bug where TT's were no longer arranged alphabetically - broken with the new layout introduced in v8.0.0. (Tks Mike.)

				v8.1.0 14th March 2020
					Restructured to support batch processing direct from the pipeline. Added new batch example:
						gci *.tar | .\Convert-RibbonSbcConfigToWord.ps1 -SkipUpdateCheck -MakePDF
					Added new bits in 8.1.0:
						- "Country Code" in System / Node Level Settings
						- "Bad Actors" table under Settings / Security
						- "Explicit Acknowledgement of Pre-Login Info" in Security / Users / Global Security Options
						- "Certificate" in Security / TLS Profiles
						- "SBC Supplementary Certificates" in Security / SBC Certificates
						- "SILK" in System / Licensing / Current Licenses
						- "DC Enabled" in Auth and Directory Services / Acive Directory / Domain Controllers
					Fixed bugs:
						- "Lockout duration" was always showing under Security / Users / Global Security Options, even if Number of Failed Logins was set to "No lockout"
						- Was not correctly handling "SupportsShouldProcess". As resolving this will add a confirm prompt for each iteration (potentially breaking
							existing workflow), I've opted to disable it
						- Under System / Node-Level Settings, 'Power LED' was misspelt
					Updated label on SIP Sig Sp 'Outbound Proxy' to now show as 'Outbound Proxy IP/FQDN'

				v8.0.3 * Not released.
					- Added display of settings when Security / Users / Global Security Options / Enhanced Password Security = Yes

				v8.0.2 * Not released.
					- No new functionality, no bugs unearthed.

				v8.0.1B 21st April 2019
					Changed Sig Gp 'Supported Video/Application Modes' to display '[None]' if $LicencedForVideo but no modes applicable (rather than blank)
					Cosmetic: moved 'Static Host' in SIP profiles between 'FQDN in From' and 'Contact' headers to align with on-screen
					Updated 'Static Host' to new label 'Static Host FQDN/IP[:port]'
					Added another example, showing how you can batch and drop the output files in another folder. (Thanks Cookie!)
					Bug CHOR-3343 in 8.0.1 b180 & 181 results in the SWeLite dropping its hostname from nodeinfo.txt, replaced by "/usr/sbin/vsbcinfo.sh: line 74: hostname: not found".
						- Updated the script to drop a warning and continue, rather than abort
					Re-jigged wording of SIP Sig Gp 'Supported Audio Modes' and 'Supported Video/Application Modes' to align with current on-screen experience
					Updated wording of Sig Gp MOH option [2] from 'Enabled for 2-way Hold Only' to 'Enabled for SDP Inactive'
					Added new bits in SweLite v8.0.0:
						- ICE Support added to SIP Sig Gps. (Removed previous test for Lite that hid this)
						- Opus codec to the media profile options
					Fixed bugs:
						- 'Static Host' now correctly displays if $FQDNinContactHeader = 'Static' [3]
						- Fixed one bad example: was using "where" shortcut instead of "foreach"
						- Corrected wrong quotes in the write-warning message under "$NodeHostname.CompareTo". Was not resolving variable names
						- SIP Sig Gp was not showing 'Nonce Lifetime' if Nonce Expiry = Limited. Deleted the now-redundant $SipNonceExpiryLookup

				v8.0.1 - 30th January 2019
					Added new bits in 8.0.1:
						- Changed ISDN Sig Gp for new 'Until Timer Expiry' option in 'Play Inband Message Post-Disconnect'
						- Added new 'Convert DISC PI=1,8 to Progress' to ISDN Sig Gp
					Fixed bugs:
						- If the SBC has only 1 Transformation Table it would not display. (Bug in the alpha sort introduced in v7.0.0B)
						- Corrected several bugs in display of 'DTLS-SRTP Profiles'
						- Corrected bug in Media Lists where the 'DTLS-SRTP Profile' value wasn't correctly reported
						- Certs was truncating parameters if one happened to be a URL (as I was splitting on '/')
						- Certs wasn't populating the issuer's details (the RHS of the table)
					- Created new $SipProfileCntctHdrLookup & edited values in $SipProfileFrmHdrLookup for new on-screen text
					- Added "[No CN]" text if a cert (so far GoDaddy roots) has no CN

				v8.0.0B - 16th November 2018
					Updated label and display of SIP Sig Gp / Media Info / Supported Audio(/Fax) Modes to match current display syntax: Lite and 1k2k
					Fixed bug:
						- The logic for displaying SIP Sig Gp ICE Mode was faulty. (Mitsu-San yet again on the ball!)

				v8.0.0 - 26th September 2018
					Updated name to "Convert-RibbonSbcConfigToWord.ps1"
					Added new bits in 8.0.0 (b502):
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
					Added new test in Test-ForNull to drop a marker in the doc if an expected lookup returned $null
					Set Word to disable spelling and grammar checking in the document. (Thanks Mike!)
					Hid the display of "Translations" (long obsolete)
					Fixed bugs:
						- Renamed variables in the SGPeers' .split() code to make them unique. (Had accidentally reused old having copy/pasted)
						- Corrected display of 'SDES-SRTP Profiles' Derivation Rate to add the '2^' prefix for non-zero values
						- If you ran with "-do calls" it would incorrectly not show the "incomplete" warning after the TOC. ("calls" sounds like "all")


				Intermediate release details have been removed from the released script
				The full revision history is here: https://greiginsydney.com/Convert-RibbonSbcConfigToWord-ps1-revision-history/


				v1.0: 2nd July 2013
					Initial release.


.LINK
	https://greiginsydney.com/uxbuilder

.EXAMPLE
	.\Convert-RibbonSbcConfigToWord.ps1

	Description
	-----------
	Defaults to a backup filename of 'symphonyconfig.xml' in the local directory, and creates a new document
	called 'symphonyconfig.docx' in the local directory. If the Word document already exists it will be overwritten
	without prompting.

.EXAMPLE
	.\Convert-RibbonSbcConfigToWord.ps1 -InputFile 'Customer-symphonyconfig.xml' -OutputFile 'CustomerUX-config.docx'

	Description
	-----------
	Opens the nominated backup file (extracted from a UX/SBC backup archive) and creates a new Word document called
	'CustomerUX-config.docx'. If 'CustomerUX-config.docx' already exists it will be overwritten without prompting.

.EXAMPLE
	.\Convert-RibbonSbcConfigToWord.ps1 -InputFile 'Customer-symphonyconfig.xml' -NodeID 'C12345678' -HardwareID '1234567890abcdef' -MakePDF -Landscape

	Description
	-----------
	Opens the nominated backup file (extracted from a UX/SBC backup archive) and creates a new Word document called
	'Customer-symphonyconfig.docx', as well as 'Customer-symphonyconfig.PDF'.
	The provided values for Serial Number and Node ID will be added to the 'Node-Level Settings' table
	If either of the output files exist they will be overwritten without prompting.
	The entire document will be created in 'landscape' layout and also saved as a PDF.

.EXAMPLE
	get-item 'd:\<path>\*.tar' | .\Convert-RibbonSbcConfigToWord.ps1 -SkipUpdateCheck -MakePDF -Verbose

	Description
	-----------
	Here's how to do batch processing. This command converts all of the .tar files in the nominated folder to '.docx' & PDF versions, with lots of on-screen progress indicators.

.EXAMPLE
	get-item 'd:\<path>\*.tar' | foreach {.\Convert-RibbonSbcConfigToWord.ps1 -InputFile $_.Fullname -OutputFile "w:\<path>\$($_.Basename).docx" -SkipUpdateCheck }

	Description
	-----------
	More batch processing. This command takes all the .tar files from d:\<path> and saves the resulting .docx files to w:\<path>.

.EXAMPLE
	Get-ChildItem symphonyconfig.xml -path c:\<Path>\ -Recurse | foreach { .\Convert-RibbonSbcConfigToWord.ps1 $_.FullName -IncludeNodeInfo -SkipUpdateCheck -MakePDF}

	Description
	-----------
	Here's how to do batch processing where you've unzipped multiple backup archives into their own directory.
	This command converts all of the 'symphonyconfig.xml' files in the nominated folder and all sub-folders to '.docx' & PDF versions.
	It also incorporates the 'NodeInfo.txt' file's content into the document if one exists in the same directory.

.EXAMPLE
	.\Convert-RibbonSbcConfigToWord.ps1 -InputFile 'Customer-symphonyconfig.xml' -Do Calls,SIP

	Description
	-----------
	Opens the nominated input file but only extracts the 'calls' (e.g. transformations, call routes, etc) and 'SIP' sections.
	Valid options are: All, IP, SIP, Sig(nalling), Calls, System, Misc & Maint. Not case-sensitive. Provide multiple options separated by a comma.
	If this switch is not specified, the default is 'All'.


.PARAMETER InputFile
		File name (and path if you wish) of the backup file. If not specified, defaults to 'symphonyconfig.xml' in the current directory.
		If specified, the InputFile must be either an .xml file, or a backup .tar.

.PARAMETER FullName
		String. File name (and path if you wish) of the backup file. This is the name of the 'InputFile' parameter when it's passed via the pipeline

.PARAMETER IncludeNodeInfo
		Optional. If the input file is NOT a .tar and you add this switch, the script will look for an open a 'nodeinfo.txt' file in the same dir as the input file.
		If the NodeInfo file isn't found the script will throw a warning and continue.

.PARAMETER OutputFile
		File name (and path if you wish) of the output file. If not specified, defaults to '<InputFile>.docx', and WILL over-write any existing file there of the same name.

.PARAMETER NodeID
		Optional. If specified, the script will embed this in the 'Node-Level Settings' table. Can also be specified as 'SerialNumber'

.PARAMETER HardwareID
		Optional. If specified, the script will embed this in the 'Node-Level Settings' table.

.PARAMETER MakePDF
		Boolean. If $True (or simply present), the created document will also be saved as a PDF, and WILL over-write any existing file there of the same name.

.PARAMETER SkipWrite
		Boolean. If $True (or simply present), the script will run but 'WriteSection' returns without doing anything. I used this as a debugging aid during script
		development, and have left it in the final version to help anyone refining or enhancing the script.

.PARAMETER Do
		Optional string. Executes only a sub-section of the script. Follow with any of All, IP, SIP, Calls, Sig, Sys, System, Misc, Maint. Separate multiple values
		with a comma. (See Examples).

.PARAMETER Landscape
		Boolean. If $True (or simply present), the script will create the entire document in Landscape view. (The default is to start in Portrait and then switch to landscape for the tables).

.PARAMETER RedactIP
		Boolean. Replaces all IP addresses in the document with dummy values.

.PARAMETER TitleColour
		String. The background colour of the title at the top of each table. Accepts a WordColor number (an Int32, e.g. '15792880') or RGB value as '240,45,27'

.PARAMETER LabelColour
		String. The background colour of each label in a table. Accepts a WordColor number (an Int32, e.g. '15792880') or RGB value as '240,45,27'

.PARAMETER SkipUpdateCheck
		Boolean. Skips the automatic check for an Update. Courtesy of Pat: http://www.ucunleashed.com/3168

.PARAMETER WdTemplate
		String. The full name and path to a Word template. Where provided, this overrides all hard-coded colour and font choices (and the optional Colour parameters)

.PARAMETER WdTableStyle
		String. The name of the Style to be applied to the tables in Word.

.PARAMETER SonusGreen
		Boolean. Writes the tables to Word in the "Sonus green" format used in versions prior to 8.0

#>

[CmdletBinding(SupportsShouldProcess = $False, DefaultParameterSetName='Legacy')]
param(
	[parameter(ParameterSetName='Legacy', ValueFromPipelineByPropertyName = $true, Position=0)]
	[alias('i')][string]$InputFile,
	[parameter(ParameterSetName='Pipeline', ValueFromPipeLine = $true, ValueFromPipelineByPropertyName = $true)]
	[string]$FullName,
	[parameter(ParameterSetName='Legacy')]
	[alias('o')][string]$OutputFile,
	[parameter(ParameterSetName='Legacy')]
	[switch]$IncludeNodeInfo,
	[parameter(ParameterSetName='Legacy')]
	[alias('serial')][string]$NodeID,
	[parameter(ParameterSetName='Legacy')]
	[alias('hw')][string]$HardwareID,
	[alias('m')][switch]$MakePDF,
	[alias('s')][switch]$SkipWrite,
	[ValidateSet('All', 'IP', 'SIP', 'Calls', 'Sig', 'Sys', 'System', 'Misc', 'Maint')]
	[String[]]
	$Do='All',
	[alias('l')][switch]$Landscape,
	[alias('r')][switch]$RedactIP,
	[alias('TitleColor')][string]$TitleColour,
	[alias('LabelColor')][string]$LabelColour,
	[switch] $SkipUpdateCheck,
	[string] $WdTemplate,
	[string] $WdTableStyle,
	[switch] $SonusGreen

)

#--------------------------------
# As of v8.1.0 the whole script lives within begin/process/end tags for enhanced pipeline support & improved efficiency (Word stays open).
#--------------------------------

begin
{

	$ScriptVersion = '9.0.2'  #Written to the title page of the document & also used in Get-UpdateInfo (as of v7.0)
	$Error.Clear()		  #Clear PowerShell's error variable
	$Global:Debug = $psboundparameters.debug.ispresent

	#--------------------------------
	# Setup hash tables--------------
	#--------------------------------

	$EnabledLookup= @{'0' = 'Disabled'; '1' = 'Enabled'; '2' = 'Drain'} #Used everywhere (Drain only used in Sig Gps)
	$ReverseEnabledLookup= @{'0' = 'Enabled'; '1' = 'Disabled'} #Seriously - Media list profiles, this is backwards!
	$EnableLookup= @{'0' = 'Disable'; '1' = 'Enable'}	#Used in multiple places. (Pedantic?)
	$TrueFalseLookup= @{'0' = 'False'; '1' = 'True'}	 #Used everywhere
	$YesNoLookup= @{'0' = 'No'; '1' = 'Yes'}			 #Used everywhere

	$FXSFXOCountryLookup = @{'1' = 'Australia'; '2' = 'Austria'; '3' = 'Belgium'; '4' = 'Brazil'; '5' = 'China'; '6' = 'Czech Republic'; '7' = 'Denmark'; '8' = 'Finland'; '9' = 'France'; '10' = 'Germany'; '11' = 'Switzerland'; '12' = 'Greece'; '13' = 'Hungary'; '14' = 'India'; '15' = 'Italy'; '16' = 'Japan'; '17' = 'Korea (ROK)'; '18' = 'Mexico'; '19' = 'Netherlands'; '20' = 'Kiwi Land'; '21' = 'Nigeria'; '22' = 'Norway'; '23' = 'Portugal'; '24' = 'Russia'; '25' = 'Saudi Arabia'; '26' = 'Slovakia'; '27' = 'South Africa'; '28' = 'Spain'; '29' = 'Sweden'; '30' = 'Taiwan'; '31' = 'Turkey'; '32' = 'United Kingdom'; '33' = 'United States'; '70' = 'United Arab Emirates'; '71' = 'Yemen'; '72' = 'Other'; '73' = 'Other European (TBR-21)'}
	$FXORingDetectLookup = @{'0' = '17'; '2' = '24'; '3' = '50'}

	$LogicalAddressIPAssignmentMethodLookup = @{'0' = 'Static'; '1' = 'DHCP'}
	$NodeLevelNetworkSettingsLookup = @{'0' = 'Static'; '1' = 'Dynamic'}

	$DCTypeLookup = @{'0' = 'Authentication'; '1' = 'Call Route'; '2' = 'On Premises'}
	$ADOperatingModeLookup = @{'0' = 'Updates'; '1' = 'Online'; '2' = 'Auth-Only'; '3' = 'Cache Lookup Only'}
	$AccountAccessLevelLookup = @{'0' = 'Audit'; '1' = 'Read-only'; '2' = 'Operator'; '3' = 'Administrator'; '4' = 'Reject User' ; '5' = 'REST'}
	$RadiusAccountingModeLookup = @{'0' = 'Active-Active'; '1' = 'Active-StandBy'; '2' = 'Round-Robin'}
	$ConsecCharsLookup = @{'0' = 'No Restriction'; '1' = '1'; '2' = '2'; '3' = '3'; '4' = '4'}
	$TxClockSourceLookup = @{'0' = 'System'; '1' = 'Network'}
	$RelayConfigStateLookup = @{'0' = 'Online'; '1' = 'Passthrough'; '2' = 'Powerdown Passthrough'}

	#Translations & Transformations
	$InputFieldLookup  = @{'0' = 'Called Address/Number'; '1' = 'Called Numbering Type'; '2' = 'Called Numbering Plan'; '3' = 'Calling Address/Number'; '4' = 'Calling Numbering Type'; '5' = 'Calling Numbering Plan'; '6' = 'Calling Number Presentation'; '7' = 'Calling Number Screening'; '8' = 'Calling Name'; '9' = 'Original Called Number'; '10' = 'Redirecting Number'; '11' = 'Redirecting Number Type'; '12' = 'Redirecting Numbering Plan'; '13' = 'Transfer Capability'; '14' = 'User Value 1'; '15' = 'User Value 2'; '16' = 'User Value 3'; '17' = 'User Value 4'; '18' = 'User Value 5'; '19' = 'Calling Number, User Specified'; '20' = 'ELIN Identifier'; '21' = 'Called Extension'; '22' = 'Calling Extension'; '23' = 'Called Phone Context'; '24' = 'Calling Phone Context'; '25' = 'Original Destination Number'; '26' = 'Callback Pool Identifier'; '27' = 'SG User Value 1'; '28' = 'SG User Value 2'; '29' = 'SG User Value 3'; '30' = 'SG User Value 4'; '31' = 'SG User Value 5'; '32' = 'Called SubAddress/Number'; '33' = 'Calling SubAddress/Number'; '34' = 'Destination Trunk Group'; '35' = 'Unknown'; '36' = 'Presence: Called'; '37' = 'Presence: Calling'; '38' = 'Called Free Phone Number'; '39' = 'Unknown'; '40' = 'Unknown'}
	$OutputFieldLookup = @{'0' = 'Called Address/Number'; '1' = 'Called Numbering Type'; '2' = 'Called Numbering Plan'; '3' = 'Calling Address/Number'; '4' = 'Calling Numbering Type'; '5' = 'Calling Numbering Plan'; '6' = 'Calling Number Presentation'; '7' = 'Calling Number Screening'; '8' = 'Calling Name'; '9' = 'Original Called Number'; '10' = 'Redirecting Number'; '11' = 'Redirecting Number Type'; '12' = 'Redirecting Numbering Plan'; '13' = 'Transfer Capability'; '14' = 'User Value 1'; '15' = 'User Value 2'; '16' = 'User Value 3'; '17' = 'User Value 4'; '18' = 'User Value 5'; '19' = 'Calling Number, User Specified'; '20' = 'ELIN Identifier'; '21' = 'Called Extension'; '22' = 'Calling Extension'; '23' = 'Called Phone Context'; '24' = 'Calling Phone Context'; '25' = 'Original Destination Number'; '26' = 'Callback Pool Identifier'; '27' = 'SG User Value 1'; '28' = 'SG User Value 2'; '29' = 'SG User Value 3'; '30' = 'SG User Value 4'; '31' = 'SG User Value 5'; '32' = 'Called SubAddress/Number'; '33' = 'Calling SubAddress/Number'; '34' = 'Destination Trunk Group'; '35' = 'Unknown'; '36' = 'Presence: Called'; '37' = 'Presence: Calling'; '38' = 'Called Free Phone Number'; '39' = 'Unknown'; '40' = 'Unknown'}
	$NumberTypeLookup = @{'-1' = 'Any/Untranslated'; '0' = 'Unknown' ; '1' = 'Subscriber'; '2' = 'National'; '3' = 'Network-Specific'; '5' = 'International'}
	$NumberingPlanLookup = @{'-1' = 'Any/Untranslated'; '0' = 'Unknown' ; '1' = 'ISDN'; '4' = 'National'; '5' = 'Private' ; '6' = 'Telephony'}
	$CallingNumberPresentationLookup = @{'-1' = 'Any/Untranslated'; '0' = 'Allowed' ; '1' = 'Restricted'; '2' = 'Not Available'; '3' = 'Reserved' }
	$CallingNumberScreeningLookup = @{'-1' = 'Any/Untranslated'; '0' = 'User - Not Screened' ; '1' = 'User - Verified'; '2' = 'User - Failed'; '3' = 'Network' }
	$TransferCapabilityLookup = @{'-1' = 'Any/Unspecified'; '0' = 'Speech'}
	$MatchTypeLookup = @{ '0' = 'Mandatory'; '1' = 'Optional'}
	$MatchTypeLookupLong = @{ '0' = "Mandatory`n(Must Match)"; '1' = "Optional`n(Match One)"}
	$TranslateActionLookup = @{'0' = '<Unknown>*'; '1' = '<Unknown>*'} #Does anyone still have a Translation table in service to tell me what these actions decode to?

	$ActionSetActionLookup = @{ '0' = 'Route Call'; '1' = 'Send Alert'; '2' = 'Send Connect'; '3' = 'Release Call'; '4' = 'Detect CNG' ; '5' = 'Route Call, Await Connect Timer'; '6' = 'Generic Timer'; '7' = 'Invoke Action Set'}
	$ActionSetExecutionLookup = @{ '0' = 'Always'; '1' = 'Prior Action Success'; '2' = 'Prior Action Failure'; '3' = 'Prior Action Timeout'}

	$ToneProfileTypeLookup = @{'0' = 'null'; '1' = 'Ringback'; '2' = 'Dial'; '3' = 'Busy'; '4' = 'Congestion'; '5' = 'Call Waiting'; '6' = 'Disconnect'; '7' = 'Confirmation'} #Tones

	$ProtocolLookup = @{'1' = 'UDP'; '2' = 'TCP'; '4' = 'TLS'}

	$LoggingProtocolLookup = @{'0' = 'UDP'; '1' = 'TCP'}
	$SysLogFacilityLookup = @{'0' = 'user (User Level Messages)'; '1' = 'kernel (Kernel Level Messages'; '2' = 'mail (Mail System) '; '3' = 'daemon (System Daemons)'; '4' = 'auth (Security/Authorization Messages)'; '5' = 'lpr (Line Printer Subsystem)'; '6' = 'news (Network News Subsystem'; '7' = 'uucp (UUCP Subsystem)'; '8' = 'cron'; '9' = 'local0 (Local Use 0)' ; '10' = 'local1 (Local Use 1)'; '11' = 'local2 (Local Use 2)'; '12' = 'local3 (Local Use 3)'; '13' = 'local4 (Local Use 4)'; '14' = 'local5 (Local Use 5)'; '15' = 'local6 (Local Use 6)'; '16' = 'local7 (Local Use 7)'}
	$LogLevelLookup = @{'0' = 'Default'; '1' = 'Trace'; '2' = 'Debug'; '3' = 'Informational'; '4' = 'Warning'; '5' = 'Error'; '6' = 'Fatal'}
	$LoggingSubsystemLookup = @{'ads' = 'Active Directory Service'; 'alarm' = 'Alarm Service'; 'mfg' = 'Analog Manufacturing Test Tool'; 'bmp' = 'BMP'; 'cas' = 'CAS Protocol'; 'route' = 'Call Routing Service'; 'csm' = 'Certificate Security Manager'; 'ccc' = 'Common Call Control'; 'trans' = 'Common Transport Service'; 'conf' = 'Configuration System'; 'zebos' = 'Core Switching/Routing'; 'dynamichosts' = 'Dynamic Hosts'; 'hosts' = 'Hosts'; 'splitdns' = 'split DNS'; 'isdn' = 'ISDN Protocol'; 'lic' = 'License Management Service'; 'log' = 'Logging System'; 'msc' = 'Media Stream Control Service'; 'symsock' = 'Network Services'; 'osys' = 'Operating System Logs'; 'pmapi' = 'Platform Management' ; 'rad' = 'RADIUS Manager Service'; 'resrc' = 'Resource Management'; 'security' = 'Security'; 'sip' = 'SIP Stack Service'; 'snmp' = 'SNMP'; 'watch' = 'Statistics & Live Info Service'; 'sba' = 'Survivable Branch Appliance'; 'shm' = 'System Health Monitoring'; 'sysio' = 'System I/O Service'; 'stmr' = 'System Timer Service'; 'tip' = 'Telephony Interface Process'; 'trapd' = 'Trap Service (SNMP)'; 'uas' = 'User Authentication Service'}

	$AclProtocolLookup = @{'1' = 'ICMP'; '6' = 'TCP'; '17' = 'UDP'; '58' = 'ICMPv6'; '89' = 'OSPF'; '256' = 'Any'} # I'll fill this with remaining 1:1 entries later
	$AclProtocolServiceLookup = @{'80' = 'HTTP'; '443' = 'HTTPS'; '22' = 'SSH'; '53' = 'DNS'; '67' = 'DHCP/BOOTP (Server)'; '161' = 'SNMP' ; '520' = 'RIP'}
	$AclAllowLookup = @{'0' = 'Allow'; '1' = 'Deny'}

	$IsdnRingbackLookup = @{'0' = 'Auto on Alert'; '1' = 'Always on Alert'; '2' = 'Never'; '3' = 'Auto on Alert/Progress'; '4' = 'Always on Alert/Progress'}
	$SipRingbackLookup = @{'0' = 'Auto on 180'; '1' = 'Always on 180'; '2' = 'Never'; '3' = 'Auto on 180/183'; '4' = 'Always on 180/183'}
	$CasRingbackLookup = @{'0' = 'Auto'; '1' = 'Always'; '2' = 'Never'; '3' = 'Auto on ???'; '4' = 'Always on ???'}
	$DirectionLookup = @{'0' = 'Inbound'; '1' = 'Outbound'; '2' = 'Bidirectional'}
	$SipSgMOHLookup = @{'0' = 'Disabled'; '1' = 'Always Enabled'; '2' = 'Enabled for SDP Inactive'}
	$SipLoadBalancingLookup = @{'0' = 'Round Robin'; '1' = 'Priority: Register All'; '2' = 'First'; '3' = '??'; '4' = 'Priority: Register Active Only'}
	$SipSgInteropModeLookup = @{'0' = 'Standard'; '1' = 'Broadsoft Extension' ; '2' = 'Office 365'; '3' = 'Office 365 w/AD PBX'; '4' = 'Header Transparency'}
	$ICEModeLookup = @{'0' = 'Lite'; '1' = 'Full'}

	$CRMediaModeLookup = @{'0' = 'DSP'; '1' = 'Proxy' ; '2' = 'Proxy preferred over DSP' ; '3' = 'DSP Preferred over Proxy' ; '4' = 'Disabled' ; '5' = 'Direct'} #New descriptions in v4
	$CRDestinationTypeLookup = @{'0' = 'Normal'; '1' = 'Registrar Table' ; '2' = 'Deny'}
	$CRCallPriorityLookup = @{'0' = 'Non-Urgent'; '1' = 'Normal'; '2' = 'Urgent'; '4' = 'Emergency'}
	$CRTODLookup = @{'0' = 'Sunday'; '1' = 'Monday' ; '2' = 'Tuesday' ; '3' = 'Wednesday' ; '4' = 'Thursday' ; '5' = 'Friday' ; '6' = 'Saturday'}

	$IsdnSideOrientationLookup = @{'0' = 'User'; '1' = 'Network'}
	$IsdnSwitchVariantLookup = @{'0' = 'ATT 4ESS'; '1' = 'ATT 5ESS'; '2' = 'DMS 100'; '3' = 'NI2'; '4' = 'Euro ISDN'; '5' = 'QSIG'; '6' = 'Japanese ISDN'}
	$SgHuntMethodLookup = @{'0' = 'Standard'; '1' = 'Reverse Standard'; '2' = 'Round Robin'; '3' = 'Least Idle'; '4' = 'Most Idle'; '5' = 'Own Number'}
	$IsdnPortPhysicalType = @{'0' = 'E1'; '1' = 'T1'}
	$IsdnPortLoopbackType = @{'0' = 'Static'; '1' = 'Line'; '2' = 'Line'; '4' = 'None'}
	$IsdnPortSignalingType = @{'0' = 'ISDN'; '1' = 'CAS'}
	$IsdnPortCodingType = @{'0' = 'B8ZS'; '1' = 'AMI'; '2' = 'HDB3'}
	$IsdnPortDS1Framing = @{'0' = 'ESF'; '1' = 'D4'}
	$IsdnIndicatedChannelLookup = @{'0' = 'Preferred'; '1' = 'Exclusive'}
	$IsdnAddPiToSetupLookup = @{'0' = 'None'; '1' = 'Not end to end ISDN'; '2' = 'Destination not ISDN'; '4' = 'Origin not ISDN'; '8' = 'Return to ISDN'; '16' = 'Interworking Encountered'; '32' = 'Inband Information'; '64' = 'Delay Encountered'}
	$IsdnNsfInfoLookup = @{'0' = 'SDN'; '1' = 'Toll Free MEGACOM'; '2' = 'MEGACOM'; '3' = 'ACCUNET Switched Digital Service'; '4' = 'International Toll Free'; '5' = 'MultiQuest'; '6' = 'Long Distance'; '7' = 'Call Redirection'}
	$IsdnNsfIdLookup = @{'0' = 'None'; '1' = 'User With ATT'; '2' = 'National With ATT'; '3' = 'International With ATT'}
	$IsdnChannelNumberBitLookup = @{'0' = 'Unset'; '1' = 'Set'}
	$IsdnAsn1ProtocolLookup = @{'0' = 'Network Extensions'; '1' = 'ROSE'}
	$IsdnAsn1NumberingLookup = @{'0' = 'Local'; '1' = 'Global'}

	$MediaRTCPModeLookup = @{'0' = 'RTCP'; '1' = 'RTCP-XR'}
	$MOHSourceLookup = @{'0' = 'File'; '1' = 'Live'}
	$CompandingLawConfigLookup = @{'0' = 'A-law'; '1' = 'u-law'}
	$EchoCancellerLookup = @{'0' = 'Standard'; '1' = 'Dual Filter'}
	$EchoCancelNLPLookup = @{'0' = 'Disabled'; '1' = 'Mild'; '2' = 'Strong'; '3' = 'Aggressive'}
	$MediaTypeLookup = @{'0' = 'G.???'; '1' = 'G.???'; '2' = 'G.711 A-Law'; '3' = 'G.711 u-Law'; '4' = '8k G.723'; '5' = 'G.726'; '6' = 'G.729'; '18' = 'G.722'; '20' = 'OPUS'; '21' = 'SILK'}
	$MediaCodecPayloadTypeLookup = @{'0' = '8k PCM u-law'; '3' = '8k GSM';'4' = '8k G.723'; '8' = '8k PCM A-law';'9' = '8k G.722';'18' = '8k G.729' } # http://www.iana.org/assignments/rtp-parameters/rtp-parameters.xml
	$MediaCryptoOperationLookup = @{'0' = '???'; '1' = 'Required'; '2' = 'Supported'; '3' = 'Off'}
	$MediaCryptoSuiteLookup = @{'0' = 'AES_CM_128_HMAC_SHA1_32'; '1' = 'AES_CM_128_HMAC_SHA1_80'}
	$DTLSCryptoSuiteLookup = @{'1' = 'AES_CM_128_HMAC_SHA1_32'; '2' = 'AES_CM_128_HMAC_SHA1_80'}
	$CodecBandwidthLookup = @{'0' = 'Narrowband'; '2' = 'Wideband'}

	$FaxTrainingConfirmationLookup = @{'0' = 'Locally Generate';'1' = 'Send Over Network'}

	$MsgXlatMsgTypeLookup = @{'0' = 'Progress / 183 Session Progress'; '1' = 'Alerting / 180 Ringing'; '2' = 'Connect / 200 OK to INVITE'; '3' = 'Proceeding (ISDN Only)'}
	$MsgXlatIETypeLookup = @{'0' = 'Progress Indicator'; '1' = "Don't Care"; '2' = 'Facility'}
	$MsgXlatOutIELookup = @{'0' = 'Untranslated (SIP/ISDN )'; '1' = 'Not Present (SIP/ISDN)'; '2' = 'Present (SIP/ISDN)'; '3' = 'Not end to end ISDN (ISDN Only)'; '4' = 'Destination not ISDN (ISDN Only)'; '5' = 'Origin not ISDN (ISDN Only)'; '6' = 'Return to ISDN (ISDN Only)'; '7' = 'Interworking Encountered (ISDN Only)'; '8' = 'Inband Information (ISDN Only)'; '9' = 'Delay Encountered (ISDN Only)'}
	$MsgXlatMediaCutThroughTypeLookup = @{'0' = 'Yes on Early Media'; '1' = 'Yes'; '2' = 'No'}
	$MsgXlatSendMessageLookup = @{'0' = 'Yes'; '1' = 'No'; '2' = 'No on Subsequent'; '3' = 'No on Cut Through'}
	$MsgXlatEarlyMediaStatusLookup = @{'0' = 'Any'; '1' = 'Negotiated'; '2' = 'Not Negotiated'}

	$SIPDescriptionLookup = @{ '100' = '100 - Trying'; '180' = '180 - Ringing'; '181' = '181 - Call Being Forwarded'; '182' = '182 - Queued'; '183' = '183 - Session Progress'; '200' = '200 - OK'; '202' = '202 - Acceptee'; '300' = '300 - Multiple Choices'; '301' = '301 - Moved Permanently'; '302' = '302 - Moved Temporarily'; '305' = '305 - Use Proxy'; '380' = '380 - Alternative Service';
		'400' = '400 - Bad Request'; '401' = '401 - Unauthorized'; '402' = '402 - Payment Required'; '403' = '403 - Forbidden'; '404' = '404 - Not Found'; '405' = '405 - Method Not Allowed'; '406' = '406 - Not Acceptable'; '407' = '407 - Proxy Authentication Required'; '408' = '408 - Request Timeout'; '409' = '409 - Conflict'; '410' = '410 - Gone'; '411' = '411 - Length Required';`
		'413' = '413 - Request Entity Too Large'; '414' = '414 - Request-URI Too Long'; '415' = '415 - Unsupported Media Type'; '416' = '416 - Unsupported URI Scheme'; '420' = '420 - Bad Extension'; '421' = '421 - Extension Required'; '423' = '423 - Interval Too Brief'; '480' = '480 - Temporarily Unavailable'; '481' = '481 - Call-Transaction Does Not Exist'; '482' = '482 - Loop Detected';
		'483' = '483 - Too Many Hops'; '484' = '484 - Address Incomplete'; '485' = '485 - Ambiguous'; '486' = '486 - Busy Here'; '487' = '487 - Request Terminated'; '488' = '488 - Not Acceptable Here'; '491' = '491 - Request Pending'; '493' = '493 - Undecipherable'; '500' = '500 - Server Internal Error'; '501' = '501 - Not Implemented'; '502' = '502 - Bad Gateway'; '503' = '503 - Service Unavailable';`
		'504' = '504 - Server Time-out'; '505' = '505 - Version Not Supported'; '513' = '513 - Message Too Large'; '580' = '580 - Precondition Failure';  '600' = '600 - Busy Everywhere'; '603' = '603 - Decline'; '604' = '604 - Does Not Exist Anywhere'; '606' = '606 - Not Acceptable'}
	$Q850DescriptionLookup = @{'1' = '1: Unallocated Number'; '2' = '2: No Route To Transit Network'; '3' = '3: No Route To Destination'; '4' = '4: Send Special Information Tone'; '5' = '5: Misdialed Trunk Prefix'; '6' = '6: Channel Unacceptable'; '7' = '7: Call Awarded in Established Channel'; '8' = '8: Preemption'; '9' = '9: Preemption - Circuit Reserved for Reuse'; '10' = '10: Normal Call Clearing';`
		'14' = '14: QoR: Ported Number'; '16' = '16: Normal Call Clearing'; '17' = '17: User Busy'; '18' = '18: No User Responding'; '19' = '19: No Answer from User (user alerted)'; '20' = '20: Subscriber Absent'; '21' = '21: Call Rejected'; '22' = '22: Number Changed'; '23 ' = '23: Redirection to New Destination'; '24 ' = '24: Call Rejected due to Feature at the Destination'; '25' = '25: Exchange Routing Error';`
		'26' = '26: Non-selected User Clearing'; '27' = '27: Destination Out of Order'; '28' = '28: Invalid Number Format (addr incomplete)'; '29' = '29: Facility Rejected'; '30' = '30: Response to STATUS INQUIRY'; '31' = '31: Normal, Unspecified'; '34' = '34: No Circuit/Channel Available'; '38' = '38: Network Out of Order'; '39' = '39: Permanent Frame Mode Connection OoS';`
		'40' = '40: Permanent Frame Mode Connection Oper'; '41' = '41: Temporary Failure'; '42' = '42: Switching Equipment Congestion'; '43' = '43: Access Information Discarded'; '44' = '44: Requested Circuit/Channel N/A'; '46' = '46: Precedence Call Blocked'; '47' = '47: Resource Unavailable, Unspecified'; '49' = '49: Quality of Service Not Available'; '50' = '50: Requested Facility Not Subscribed';`
		'53' = '53: Outgoing Calls Barred Within CUG'; '55' = '55: Incoming Calls Barred Within CUG'; '57' = '57: Bearer Capability Not Authorized'; '58' = '58: Bearer Capability Not Available'; '62' = '62: Inconsistency in Outgoing IE'; '63' = '63: Service or Option N/A, unspecified'; '65' = '65: Bearer Capability Not Implemented'; '66' = '66: Channel Type Not Implemented'; '69' = '69: Requested Facility Not Implemented';`
		'70' = '70: Only Restricted Digital Bearer Cap Supported'; '79' = '79: Service or Option not Implemented, Unspecified'; '81' = '81: Invalid Call Reference Value'; '82' = '82: Identified Channel Does Not Exist'; '83' = '83: Call Exists, but Call Identity Does Not'; '84' = '84: Call Identity in User'; '85' = '85: No Call Suspended'; '86' = '86: Call eith Requested Call Identity has Cleared';`
		'87' = '87: User Not Member of CUG'; '88' = '88: Incompatible Destination';	'90' = '90: Non-Existent CUG'; '91' = '91: Invalid Transit Network Selection'; '95' = '95: Invalid Message, Unspecified'; '96' = '96: Mandatory Information Element is Missing'; '97' = '97: Message Type Non-existent / Not Implemented'; '98' = '98: Message Incompatible With Call State or Message Type';`
		'99' = '99: IE/Parameter Non-existent or Not Implemented'; '100' = '100: Invalid Information Element Contents'; '101' = '101: Message Not Compatible with Call State'; '102' = '102: Recovery on Timer Expiry'; '103' = '103: Parameter Non-existent / Not Implemented, Passed On'; '110' = '110: Message With Unrecognized Parameter, Discarded'; '111' = '111: Protocol Error, Unspecified'; '127' = '127: Interworking, Unspecified'}

	$SIPCondRuleOperationLookup = @{'0' = 'Regex'; '1' = 'Equals'}
	$SIPCondRuleMatchTypeLookup = @{'0' = 'Literal'; '1' = 'Token'}
	$SIPMsgRuleTableActionLookup = @{'0' = 'ukwn0'; '1' = 'ukwn1'; '2' = 'ukwn2'; '3' = 'ukwn3'; '4' = 'ukwn4'}
	$SIPMsgRuleStatusLineLookup = @{'0' = 'Modify'; '1' = 'Copy Value to'; '2' = 'Ignore'}
	$SIPMsgRuleHdrOrdnlLookup = @{'0' = 'All'; '1' = '1st'; '2' = '2nd'; '3' = '3rd'; '4' = '4th'; '5' = '5th'; '6' = '6th'; '7' = '7th'; '8' = '8th'; '9' = '9th'; '10' = '10th'; '11' = '11th'; '12' = '12th'; '13' = '13th'; '14' = '14th'; '15' = '15th'; '16' = '16th'; '17' = '17th'; '18' = '18th'; '19' = '19th' ; '20' = '20th' ; '-1' = 'Last'; '-2' = '2nd from Last'; '-3' = '3rd from Last'; '-4' = '4th from Last'; '-5' = '5th from Last'; '-6' = '6th from Last'; '-7' = '7th from Last'; '-8' = '8th from Last'; '-9' = '9th from Last'; '-10' = '10th from Last'; '-11' = '11th from Last'; '-12' = '12th from Last'; '-13' = '13th from Last'; '-14' = '14th from Last'; '-15' = '15th from Last'; '-16' = '16th from Last'; '-17' = '17th from Last'; '-18' = '18th from Last'; '-19' = '19th from Last' ; '-20' = '20th from Last'}

	$CasR2SignalingTypeLookup = @{'0' = 'DTMF'; '1' = 'MF'}
	$CasR2CDBitsLookup = @{'0' = '00'; '1' = '01'; '2' = '10'; '3' = '11'}
	$CasR2InvertedABCDBitsLookup = @{'0' = '0000'; '1' = '0001'; '2' = '0010'; '3' = '0011'; '4' = '0100'; '5' = '0101' ; '6' = '0110'; '7' = '0111'; '8' = '1000'; '9' = '1001'; '10' = '1010'; '11' = '1011'; '12' = '1100'; '13' = '1101'; '14' = '1110'; '15' = '1111'}
	$CasR2DigitsRxdLookup = @{'0' = 'All'; '1' = '1'; '2' = '2'; '3' = '3'; '4' = '4'; '5' = '5' ; '6' = '6'; '7' = '7'; '8' = '8'; '9' = '9'; '10' = '10'}
	$CasR2Group1Lookup = @{'0' = 'None'; '1' = 'I-1'; '2' = 'I-2'; '3' = 'I-3'; '4' = 'I-4'; '5' = 'I-5'; '6' = 'I-6'; '7' = 'I-7'; '8' = 'I-8'; '9' = 'I-9'; '10' = 'I-10'; '11' = 'I-11'; '12' = 'I-12'; '13' = 'I-13'; '14' = 'I-14'; '15' = 'I-15'}
	$CasR2CategoryLookup = @{'0' = 'None'; '1' = 'II-1'; '2' = 'II-2'; '3' = 'II-3'; '4' = 'II-4'; '5' = 'II-5'; '6' = 'II-6'; '7' = 'II-7'; '8' = 'II-8'; '9' = 'II-9'; '10' = 'II-10'; '11' = 'II-11'; '12' = 'II-12'; '13' = 'II-13'; '14' = 'II-14'; '15' = 'II-15'}
	$CasR2GroupASignalsLookup = @{'0' = 'None'; '1' = 'A-1'; '2' = 'A-2'; '3' = 'A-3'; '4' = 'A-4'; '5' = 'A-5'; '6' = 'A-6'; '7' = 'A-7'; '8' = 'A-8'; '9' = 'A-9'; '10' = 'A-10'; '11' = 'A-11'; '12' = 'A-12'; '13' = 'A-13'; '14' = 'A-14'; '15' = 'A-15'}
	$CasR2GroupBSignalsLookup = @{'0' = 'None??'; '1' = 'B-1'; '2' = 'B-2'; '3' = 'B-3'; '4' = 'B-4'; '5' = 'B-5'; '6' = 'B-6'; '7' = 'B-7'; '8' = 'B-8'}
	$CasR2GroupBSignalsCheckLookup = @{'0' = ''; '1' = '<Checked>'}

	$CasOrientationLookup = @{'0' = 'User/Slave'; '1' = 'Network/Master'}
	$CasLoopStartTypeLookup = @{'0' = 'Basic'; '1' = 'Reverse Battery'; '2' = 'Forward Disconnect'}
	$CasStartDialLookup = @{'0' = 'Wink Start'; '1' = 'Immediate Dialing'}
	$CasSgLineTypeLookup = @{'0' = 'Analog'; '1' = 'Digital'}
	$CasSgCallerIDTypeLookup = @{'0' = 'Disabled'; '1' = 'FSK'; '2' = 'DTMF'; '3' = 'DTMF, Reverse Battery'; '4' = 'ETSI, Reverse Battery'; '5' = 'NTT, Japan CID' ; '6' = 'ETSI'; '7' = 'DTMF, No Ring'; '8' = 'FSK, No Ring'}
	$CasSgDigitalCallerIDTypeLookup = @{'0' = 'Disabled'; '1' = 'FSK'; '2' = 'DTMF'}
	$CasSgDTMFCallerIDDelimiterLookup = @{'10' = 'A'; '11' = 'B'; '12' = 'C'; '13' = 'D'; '14' = '#'; '15' = '*'}

	$LoggerLogLevelLookup = @{'0' = 'Default'; '1' = 'Trace'; '2' = 'Debug';'3' = 'Informational'; '4' = 'Warning' ; '5' = 'Error'; '6' = 'Fatal'} # (Not used in this version)

	$PortHardwareTypeLookup = @{'0' = 'Ethernet'; '1' = '<Unhandled Value>'; '2' = '<Unhandled Value>'}
	$NetworkAdapterFrameTypeLookup = @{'0' = 'All'; '1' = 'Untagged'; '2' = 'Tagged'}
	$NetworkAdapterGigabitTimingLookup = @{'0' = 'Auto (default)'; '1' = 'Master'; '2' = 'Slave'}
	$PortDuplexityLookup = @{'0' = 'Half'; '1' = 'Full'; '2' = 'Auto'}
	$PortSpeedLookup = @{'0' = '10'; '1' = '100'; '2' = '1000'; '3' = 'Auto'}
	$IpVersionLookup = @{'0' = 'IPv4'; '1' = 'IPv6'; '2' = 'Both'}
	$DHCPOptionsToUseLookup = @{'0' = 'All'; '1' = 'IP Address Only'; '2' = 'IP Address and DNS'}

	$PortMirrorDirectionLookup = @{'0' = 'Transmit'; '1' = 'Receive'; '2' = 'Transmit & Receive'}
	$PortMirrorPortLookup = @{'tap0' = 'DSP'; 'tap1' = 'CPU 1'; 'tap2' = 'CPU 2'; 'tap3' = 'tap3'; 'tap4' = 'tap4'; 'tap5' = 'tap5'; 'tap6' = 'tap6'; 'tap7' = 'tap7'; 'tap8' = 'tap8'; 'tap9' = 'tap9'; 'tap10' = 'tap10'; 'tap11' = 'tap11'; 'tap12' = 'tap12'; 'tap13' = 'tap13'; 'tap14' = 'Ethernet 1'; 'tap15' = 'Ethernet 2'; 'tap16' = 'Ethernet 3'; 'tap17' = 'Ethernet 4'; 'tap18' = 'tap18'; 'tap19' = 'ASM 2' ; 'tap20' = 'ASM 1'}

	$MstpInstanceBridgePriorityLookup = @{'0' = '0'; '1' = '4096'; '2' = '8192'; '3' = '12288'; '4' = '16384'; '5' = '20480' ; '6' = '24576'; '7' = '28672'; '8' = '32768'; '9' = '36864'; '10' = '40960'; '11' = '45056'; '12' = '49152'; '13' = '53248'; '14' = '57344'; '15' = '61440'}
	$BridgeRegionSettingsProtocolLookup = @{'0' = 'MSTP'; '1' = '<Unhandled Value>'; '2' = '<Unhandled Value>' }

	$LinkMonitorHostTypeLookup = @{'0' = 'Host'; '1' = 'Gateway' }
	$LinkMonitorTypeLookup = @{'0' = 'CAC/IPSEC'; '1' = 'Backup Default Route' }

	#SIP Profiles
	$SipProfileFrmHdrLookup = @{'0' = 'Disable'; '1' = 'SBC Edge FQDN'; '2' = 'Server FQDN'; '3' = 'Static'}
	$SipProfileCntctHdrLookup = @{'0' = 'Disable'; '1' = 'SBC FQDN'; '2' = '<Unhandled Value>'; '3' = 'Static'}
	$SipProfileRefreshLookup = @{'0' = 'Reinvite'; '1' = 'Update'}
	$SipProfileElinIDLookup = @{'0' = 'LOC'; '1' = 'HNO'; '2' = 'FLR'}
	$SipProfileAssertHdrLookup = @{'0' = 'Trusted Only'; '1' = 'Always'; '2' = 'Never'}
	$SipProfileOptionsLookup = @{'0' = '<n/a>'; '1' = 'Supported'; '2' = 'Required'; '3' = 'Not Present'}
	$SipProfileDivHrdLookup = @{'0' = 'Last'; '1' = 'First'}
	$SipProfileRecordHrdLookup = @{'0' = 'RFC 3261 Standard'; '1' = 'ETSI Standard'}

	$SipProfileClgInfoSourceLookup = @{'0' = 'RFC Standard'; '1' = "'From' Header Only"}
	$SipProfileMaxReTxLookup = @{'0' = 'RFC Standard'; '1' = '1'; '2' = '2'; '3' = '3'; '4' = '4'; '5' = '5'; '6' = '6'; '7' = '7'; '8' = '8'; '9' = '9'; '10' = '10'}
	$SipProfileDigitPrefLookup = @{'0' = 'SIP INFO'; '1' = 'RFC 2833/Voice' }
	$SipProfileSDPHandlingLookup = @{'0' = 'Legacy Audio/Fax'; '1' = 'RFC 3264' }

	$SipElementDescActionLookup = @{'0' = '0?'; '1' = 'Add'; '2' = 'Modify'; '3' = 'Remove'; '4' = 'Copy Value To'; '5' = 'Ignore'}
	$SipElementDescElementClassLookup = @{'0' = '<Default>'; '1' = 'Display Name'; '2' = 'URI'; '3' = 'URI Scheme'; '4' = 'URI Host'; '5' = 'URI User Info'; '6' = 'URI User'; '7' = 'Password'; '8' = 'URI Port'; '9' = 'URI/Header Parameters'; '10' = 'Method'; '11' = 'SIP Version'; '12' = 'Status'}

	$AORTypeLookup = @{'0' = '<Invalid?>'; '1' = 'Local'; '2' = 'Remote'; '3' = 'Static'}
	$RemoteAuthFromURILookup = @{'0' = 'Authentication ID'; '1' = 'Regex'}

	$TlsClientCipherLookup = @{'1' = 'AES128-SHA'; '2' = 'DES-CBC3-SHA'; '3' = 'AES128-SHA, DES-CBC3-SHA'; '4' = 'DES-CBC-SHA'; '5' = 'AES128-SHA, DES-CBC3-SHA, DES-CBC-SHA'}
	$TlsClientCipherLookupV4 = @{'0' = 'TLS_RSA_WITH_AES128_CBC_SHA'; '1' = 'TLS_RSA_WITH_AES256_CBC_SHA'; '2' = 'TLS_RSA_WITH_3DES_EDE_CBC_SHA'; '3' = 'TLS_RSA_WITH_AES_128_CBC_SHA256'; '4' = 'TLS_RSA_WITH_AES_256_CBC_SHA256'; '5' = 'TLS_ECDHE_RSA_WITH_AES_128_CBC_SHA256'; '6' = 'TLS_ECDHE_RSA_WITH_AES_256_CBC_SHA384'; '7' = 'TLS_ECDHE_RSA_WITH_3DES_EDE_CBC_SHA'}
	$TlsProtocolLookup = @{'0' = 'TLS 1.2 Only'; '1' = 'TLS 1.0 Only'; '2' = 'TLS 1.0-1.2'}
	$DTLSHashTypeLookup = @{'1' = 'DTLS_MEDIA_CRYPT0_HASH_SHA1'; '2' = 'DTLS_MEDIA_CRYPTO_HASH_SHA224'; '3' = 'DTLS_MEDIA_CRYPTO_HASH_SHA256'; '4' = 'DTLS_MEDIA_CRYPTO_HASH_SHA384'; '5' = 'DTLS_MEDIA_CRYPTO_HASH_SHA512'; '6' = 'DTLS_MEDIA_CRYPT0_HASH_MD5';}
	$BadActorTypeLookup = @{'0' = 'Calling Number'; '1' = 'Called Number'; '2' = 'IPv4 Address'; '3' = 'IPv6 Address'}

	$SIPUserInfoDecodeLookup = @{'0' = 'Legacy'; '1' = 'RFC 3261'}

	$SNMPCommunityTypeLookup = @{'0' = '<Invalid?>'; '1' = 'Read-Only'}
	$SNMPTrapTypeLookup = @{'1' = 'Alarm'; '2' = 'Event'}
	$SNMPSeverityLookup = @{'0' = 'None'; '1' = 'Warning'; '2' = 'Minor'; '3' = 'Major'; '4' = 'Critical'}
	$SNMPCategoryLookup = @{'0' = 'None'; '1' = 'Communication'; '2' = 'Equipment'; '3' = 'Processing' ; '4' = 'General'; '5' = 'Environmental'; '6' = 'QOS'; '7' = 'Security'}

	$TCAMonitoredStatisticLookup = @{'0' = '<Unhandled Value>'; '1' = 'TDM Signaling Group Channel Usage'; '2' = 'SIP Call License Usage'; '3' = 'SIP Registrations' ; '4' = 'DSP Usage'; '5' = 'CPU Usage'; '6' = 'Memory Usage'; '7' = 'File Descriptor Usage'; '8' = '1 Minute Load Average'; '9' = '5 Minute Load Average'; '10' = '15 Minute Load Average'; '11' = 'Temporary Partition Usage'; '12' = 'Logging Partition Usage'; '13' = '<Unhandled Value>'; '14' = '<Unhandled Value>'; '15' = '<Unhandled Value>'; '16' = '<Unhandled Value>'; '17' = '<Unhandled Value>'}
	$TCAMonitoredStatisticValueLookup = @{'0' = ''; '1' = ''; '2' = ''; '3' = '' ; '4' = '%'; '5' = '%'; '6' = '%'; '7' = ''; '8' = ''; '9' = ''; '10' = ''; '11' = '%'; '12' = ''; '13' = ''; '14' = ''; '15' = ''; '16' = ''; '17' = ''}

	$NotificationProviderLookup = @{'0' = 'Kandy'; '1' = '<Unhandled Value>'}
	$NotificationEventsLookup = @{'0' = 'E911'; '1' = '<Unhandled Value>'}

	$CountryCodeLookup = @{'0' = 'None'; '93' = 'Afghanistan'; '355' = 'Albania'; '213' = 'Algeria'; '376' = 'Andorra'; '244' = 'Angola'; '672' = 'Antarctica'; '54' = 'Argentina'; '374' = 'Armenia'; '297' = 'Aruba'; '61' = 'Australia'; '43' = 'Austria'; '994' = 'Azerbaijan'; '973' = 'Bahrain'; '880' = 'Bangladesh'; '375' = 'Belarus'; '32' = 'Belgium'; '501' = 'Belize'; '229' = 'Benin';`
		'975' = 'Bhutan'; '591' = 'Bolivia'; '387' = 'Bosnia and Herzegovina'; '267' = 'Botswana'; '55' = 'Brazil'; '246' = 'British Indian Ocean Territory'; '673' = 'Brunei'; '359' = 'Bulgaria'; '226' = 'Burkina Faso'; '257' = 'Burundi'; '855' = 'Cambodia'; '237' = 'Cameroon'; '1c' = 'Canada'; '238' = 'Cape Verde'; '236' = 'Central African Republic'; '235' = 'Chad'; '56' = 'Chile'; '86' = 'China';`
		'57' = 'Colombia'; '269' = 'Comoros'; '682' = 'Cook Islands'; '506' = 'Costa Rica'; '385' = 'Croatia'; '53' = 'Cuba'; '357' = 'Cyprus'; '420' = 'Czech Republic'; '243' = 'Democratic Republic of the Congo'; '45' = 'Denmark'; '253' = 'Djibouti'; '670' = 'East Timor'; '593' = 'Ecuador'; '20' = 'Egypt'; '503' = 'El Salvador'; '240' = 'Equatorial Guinea'; '291' = 'Eritrea'; '372' = 'Estonia';`
		'251' = 'Ethiopia'; '500' = 'Falkland Islands'; '298' = 'Faroe Islands'; '679' = 'Fiji'; '358' = 'Finland'; '33' = 'France'; '689' = 'French Polynesia'; '241' = 'Gabon'; '220' = 'Gambia'; '995' = 'Georgia'; '49' = 'Germany'; '233' = 'Ghana'; '350' = 'Gibraltar'; '30' = 'Greece'; '299' = 'Greenland'; '1g' = 'Guam'; '502' = 'Guatemala'; '224' = 'Guinea'; '245' = 'Guinea-Bissau'; '592' = 'Guyana';`
		'509' = 'Haiti'; '504' = 'Honduras'; '852' = 'Hong Kong'; '36' = 'Hungary'; '354' = 'Iceland'; '91' = 'India'; '62' = 'Indonesia'; '98' = 'Iran'; '964' = 'Iraq'; '353' = 'Ireland'; '972' = 'Israel'; '39' = 'Italy'; '225' = 'Ivory Coast'; '81' = 'Japan'; '962' = 'Jordan'; '7k' = 'Kazakhstan'; '254' = 'Kenya'; '686' = 'Kiribati'; '82' = 'Korea (ROK)'; '383' = 'Kosovo'; '965' = 'Kuwait';`
		'996' = 'Kyrgyzstan'; '856' = 'Laos'; '371' = 'Latvia'; '961' = 'Lebanon'; '266' = 'Lesotho'; '231' = 'Liberia'; '218' = 'Libya'; '423' = 'Liechtenstein'; '370' = 'Lithuania'; '352' = 'Luxembourg'; '853' = 'Macau'; '389' = 'Macedonia'; '261' = 'Madagascar'; '265' = 'Malawi'; '60' = 'Malaysia'; '960' = 'Maldives'; '223' = 'Mali'; '356' = 'Malta'; '692' = 'Marshall Islands'; '222' = 'Mauritania';`
		'230' = 'Mauritius'; '262' = 'Mayotte'; '52' = 'Mexico'; '691' = 'Micronesia'; '373' = 'Moldova'; '377' = 'Monaco'; '976' = 'Mongolia'; '382' = 'Montenegro'; '212' = 'Morocco'; '258' = 'Mozambique'; '95' = 'Myanmar'; '264' = 'Namibia'; '674' = 'Nauru'; '977' = 'Nepal'; '31' = 'Netherlands'; '599' = 'Netherlands Antilles'; '687' = 'New Caledonia'; '64' = 'New Zealand'; '505' = 'Nicaragua';`
		'227' = 'Niger'; '234' = 'Nigeria'; '683' = 'Niue'; '850' = 'North Korea'; '47' = 'Norway'; '968' = 'Oman'; '92' = 'Pakistan'; '680' = 'Palau'; '970' = 'Palestinian Territory'; '507' = 'Panama'; '675' = 'Papua New Guinea'; '595' = 'Paraguay'; '51' = 'Peru'; '63' = 'Philippines'; '48' = 'Poland'; '351' = 'Portugal'; '974' = 'Qatar'; '242' = 'Republic of the Congo'; '40' = 'Romania';`
		'7' = 'Russia'; '250' = 'Rwanda'; '290' = 'Saint Helena'; '508' = 'Saint Pierre and Miquelon'; '685' = 'Samoa'; '378' = 'San Marino'; '239' = 'Sao Tome and Principe'; '966' = 'Saudi Arabia'; '221' = 'Senegal'; '381' = 'Serbia'; '248' = 'Seychelles'; '232' = 'Sierra Leone'; '65' = 'Singapore'; '421' = 'Slovakia'; '386' = 'Slovenia'; '677' = 'Solomon Islands'; '252' = 'Somalia';`
		'27' = 'South Africa'; '211' = 'South Sudan'; '34' = 'Spain'; '94' = 'Sri Lanka'; '249' = 'Sudan'; '597' = 'Suriname'; '268' = 'Swaziland'; '46' = 'Sweden'; '41' = 'Switzerland'; '963' = 'Syria'; '886' = 'Taiwan'; '992' = 'Tajikistan'; '255' = 'Tanzania'; '66' = 'Thailand'; '228' = 'Togo'; '690' = 'Tokelau'; '676' = 'Tonga'; '216' = 'Tunisia'; '90' = 'Turkey'; '993' = 'Turkmenistan';`
		'688' = 'Tuvalu'; '256' = 'Uganda'; '380' = 'Ukraine'; '971' = 'United Arab Emirates'; '44' = 'United Kingdom'; '1' = 'United States'; '598' = 'Uruguay'; '998' = 'Uzbekistan'; '678' = 'Vanuatu'; '379' = 'Vatican'; '58' = 'Venezuela'; '84' = 'Vietnam'; '681' = 'Wallis and Futuna'; '967' = 'Yemen'; '260' = 'Zambia'; '263' = 'Zimbabwe'}


	#--------------------------------
	# WORD CONSTANTS ----------------
	#--------------------------------

	#https://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	$wdNormalTemplate = 0
	$wdNewBlankDocument = 0
	$wdSeekPrimaryFooter = 4
	$wdSeekMainDocument = 0
	$wdAlignPageNumberRight = 2
	$wdOrientLandscape = 1
	$wdOrientPortrait = 0
	$wdSectionBreakNextPage = 2
	$wdPageBreak = 7
	$wdAlignParagraphCenter = 1
	$wdReplaceAll = 2
	$wdReplaceNone = 0

	$wdDoNotSaveChanges = 0
	$wdMissingValue = [Reflection.Missing]::Value

	$wdStyleNormal   = -1
	$wdStyleHeading1 = -2
	$wdStyleHeading2 = -3
	$wdStyleHeading3 = -4

	# Colours: http://msdn.microsoft.com/en-us/library/office/bb237558(v=office.12).aspx
	$wdColorGray05 = 15987699
	$wdColorGray15 = 14277081
	$wdColorGray30 = 11776947
	$wdColorGray70 = 5000268
	$wdColorAutomatic = -16777216
	$wdColorPaleBlue = 16764057
	$wdColorSkyBlue = 16763904
	$wdColorLightBlue = 16737843
	$wdColorBlueGray = 10053222
	$wdColorIndigo = 10040115
	$wdColorGreen = 32768
	$wdColorLightGreen = 13434828
	$wdColorOliveGreen = 13107
	$wdColorLime = 52377
	$wdColorSeaGreen = 6723891

	$beLightGreen = 15792880	#Thanks Ben for the colour!


	#--------------------------------
	# START FUNCTIONS ---------------
	#--------------------------------

	function Write-Section ()
	{
		param ( [string]$progressBarTitle, [array]$headers, [array]$values)
		if ($SkipWrite) { return } # Skip writing to Word.
								   # I used this when debugging or adding a new section. I'd add a 'write-host' line in the XML decoding to dump the newly-tested section to the screen.

		#Create the table and add the title row
		$selection.Style=$wdStyleNormal
		$rows=1
		$columns = $headers.length
		$range = $selection.Range
		$table=$doc.Tables.add($range,$rows,$columns)
		$table.Borders.Enable=$true
		for ($i = 1; $i -le $columns; $i++)
		{
			if ($progressBarTitle -ne '')
			{
				$progressPercent2 = $i / (($values.Count+1) * $columns) * 100
				If ($progressPercent2 -gt 100) {$progressPercent2 = 100 } #Trap any accidental percentage overrun
				write-progress -id 2 -parentid 1 -Activity 'Table' -Status $($progressBarTitle) -PercentComplete $progressPercent2
			}
			switch ($headers[$i-1])
			{
				'Enabled'	 { $table.Cell(1,$i).SetWidth(50,1)}	#Constrain
				'Entry'	   { $table.Cell(1,$i).SetWidth(40,1)}		#Constrain
				'Index'	   { $table.Cell(1,$i).SetWidth(40,1)}		#Constrain
				'Primary Key' { $table.Cell(1,$i).SetWidth(60,1)}	#Constrain			- aka 'Entry' and 'Index'
				'Instance ID' { $table.Cell(1,$i).SetWidth(60,1)}	#Constrain			- Node Interfaces / Bridge
				'VLAN ID'	 { $table.Cell(1,$i).SetWidth(60,1)}	#Constrain			- Node Interfaces / Bridge
				'Card'		{ $table.Cell(1,$i).SetWidth(30,1)}		#Constrain
				'Port'		{ $table.Cell(1,$i).SetWidth(30,1)}		#Constrain
				'Port Name'   { $table.Cell(1,$i).SetWidth(60,1)}	#Constrain			- ISDN Sig Gp
				'Port ID'	 { $table.Cell(1,$i).SetWidth(50,1)}	#Constrain			- Node Interfaces / Ports
				'Rx Gain'	 { $table.Cell(1,$i).SetWidth(35,1)}	#Constrain			- FXS & FXO Ports
				'Tx Gain'	 { $table.Cell(1,$i).SetWidth(35,1)}	#Constrain			- FXS & FXO Ports
				'Port Type'	 { $table.Cell(1,$i).SetWidth(30,1)}	#Constrain			- Ports
				'Analog Line Profile'   { $table.Cell(1,$i).SetWidth(60,1)}  #Constrain	- FXS & FXO Ports
				'Max Freq'	{ $table.Cell(1,$i).SetWidth(35,1)}		#Constrain			- FXO Ports
				'Min Freq'	{ $table.Cell(1,$i).SetWidth(35,1)}		#Constrain			- FXO Ports
				'Ring Validation'	{ $table.Cell(1,$i).SetWidth(50,1)}  #Constrain	- FXO Ports
				'Ring Det Thresh'	{ $table.Cell(1,$i).SetWidth(45,1)}  #Constrain	- FXO Ports
				'Hold' 		  { $table.Cell(1,$i).SetWidth(40,1)}	#Constrain			- CAS Supp Service Profiles
				'Transfer'	{ $table.Cell(1,$i).SetWidth(40,1)}		#Constrain			- CAS Supp Service Profiles
				'Call Waiting'{ $table.Cell(1,$i).SetWidth(40,1)}	#Constrain			- CAS Supp Service Profiles
				'Conference'  { $table.Cell(1,$i).SetWidth(40,1)}	#Constrain			- CAS Supp Service Profiles
				'Priority'	{ $table.Cell(1,$i).SetWidth(40,1)}		#Constrain
				'Protocol'	{ $table.Cell(1,$i).SetWidth(50,1)}		#Constrain
				'Channels'	{ $table.Cell(1,$i).SetWidth(50,1)}		#Constrain
				'RingBack'	{ $table.Cell(1,$i).SetWidth(50,1)}		#Constrain
				'Early 183'	{ $table.Cell(1,$i).SetWidth(55,1)}	#Constrain
				'Line Type'	{ $table.Cell(1,$i).SetWidth(50,1)}	#Constrain
				#'Match Type'  { $table.Cell(1,$i).SetWidth(55,1)}	#Constrain			- Conflict: Transformation Tables & Message Manipulation / Condition Rules Tables
				'Media DSCP'{ $table.Cell(1,$i).SetWidth(35,1)}	#Constrain			- Media List
				'RTCP Mode'	{ $table.Cell(1,$i).SetWidth(40,1)}	#Constrain			- Media List
				'Dead Call Detection'   { $table.Cell(1,$i).SetWidth(45,1)}  #Constrain	- Media List
				'Silence Suppression'   { $table.Cell(1,$i).SetWidth(50,1)}  #Constrain	- Media List
				'Codec'		{ $table.Cell(1,$i).SetWidth(50,1)} 		#Constrain - Media Profiles
				'Rate  (b/s)' { $table.Cell(1,$i).SetWidth(70,1)} 	#Constrain - Media Profiles
				'Payload Size (ms)' { $table.Cell(1,$i).SetWidth(60,1)} #Constrain - Media Profiles
				'Payload Type' { $table.Cell(1,$i).SetWidth(50,1)} 	#Constrain - Media Profiles
				'Fork Call'	{ $table.Cell(1,$i).SetWidth(40,1)}  #Constrain			- Call Routes
				'Destination Type'   { $table.Cell(1,$i).SetWidth(80,1)}  #Constrain	- Call Routes
				'Rule Type'	{ $table.Cell(1,$i).SetWidth(80,1)}  #Constrain 			- Message Rule Tables
				'Result Type'  { $table.Cell(1,$i).SetWidth(80,1)}  #Constrain 			- Message Rule Tables
				'Condition Expression' { $table.Cell(1,$i).SetWidth(80,1)}  #Constrain - Message Rule Tables
				'Description' { $table.Cell(1,$i).SetWidth(120,1)} #Expand
				'Host'		{ $table.Cell(1,$i).SetWidth(120,1)} #Expand
				'Expiry'	{ $table.Cell(1,$i).SetWidth(100,1)} #Expand				- Certificates
				#'Action'	{ $table.Cell(1,$i).SetWidth(40,1)}  #Constrain			- Conflict: Access Control Lists & Action Set Tables
				'Interface Name'	{ $table.Cell(1,$i).SetWidth(55,1)}  #Constrain	- Access Control Lists
				'Precedence'{ $table.Cell(1,$i).SetWidth(50,1)}  #Constrain			- Access Control Lists
				'ID'		{ $table.Cell(1,$i).SetWidth(40,1)}  #Constrain			- SNMP Alarm/Events
				'Type'		{ $table.Cell(1,$i).SetWidth(40,1)}  #Constrain			- SNMP Alarm/Events
				'Severity'	{ $table.Cell(1,$i).SetWidth(80,1)}  #Constrain			- SNMP Alarm/Events
				'Category'	{ $table.Cell(1,$i).SetWidth(80,1)}  #Constrain			- SNMP Alarm/Events

			}
			$table.cell(1,$i).Range.text = ($headers[$i-1]).ToString() #Added as precautionary for Win10
		}
		#Add one or more lines of content
		$row=2
		$cellCount = $columns
		foreach ($value in $values)
		{
			$null = $table.rows.Add()
			for ($i = 1; $i -le $columns; $i++)
			{
				if ($progressBarTitle -ne '')
				{
					$cellCount ++
					$progressPercent2 = ($cellCount) / (($values.Count+1) * $columns) * 100
					If ($progressPercent2 -gt 100) { $progressPercent2 = 100 } #Trap any accidental percentage overrun
					write-progress -id 2 -parentid 1 -Activity 'Table' -Status $($progressBarTitle) -PercentComplete $progressPercent2
				}
				if ($value[$i-1] -eq $null) { continue }
				$table.cell($row,$i).Range.text = ($value[$i-1]).ToString() #Added as precautionary for Win10
			}
			$row ++
		}
		if ($row -eq 2)
		{
			#We didn't increment the row counter after it was initialised - this table is empty
			$null = $table.rows.Add()
			$table.cell($row,1).Range.text = '-- Table is empty --'
			$table.cell($row,1).Merge($table.cell($row,$columns))
			$table.cell($row,1).Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
		}
		# Bold the header row:
		$table.Rows.First.HeadingFormat = -1
		$table.Rows.First.Range.Font.Bold = $true
		if (!$wdTemplate) { $table.Rows.First.Shading.BackgroundPatternColor = $iTitleColour } # $wdColorFirstRow
		if ($progressBarTitle -ne '')
		{
			write-progress -id 2 -parentid 1 -Activity 'Table' -Status $($progressBarTitle) -PercentComplete (100) -completed
		}
		#Step out of the table:
		$null = $Selection.EndKey(6, 0)
		$selection.TypeParagraph() #Stick a blank line after each table
	}


	function Write-SectionVertically ()
	{
		param ( [string]$progressBarTitle, [array]$values)

		if ($SkipWrite) { return } # Skip writing to Word.
								   # I used this when debugging or adding a new section. I'd add a 'write-host' line in the XML decoding to dump the newly-tested section to the screen.

		#Create the table (no title row for vertical)
		$selection.Style=$wdStyleNormal
		$range = $selection.Range
		$table=$doc.Tables.add($range,1,5)

		#Constrain the column widths
		$table.Cell(1,3).SetWidth(20,1)
		$table.cell(1,3).Range.text = ''	#Setting the width will crash if we've not yet written any values to that cell/column
		$table.Cell(1,1).SetWidth(160,1)
		$table.Cell(1,4).SetWidth(160,1)
		#Add the rest of the rows - they'll now follow the Width from the first
		for ($i=1; $i -lt  $values.Count; $i++)
		{
			$null = $table.rows.Add()
		}

		$table.Borders.Enable=$true
		#Add one or more lines of content
		$row = 1
		$column = 0
		$cellCount = $values.Count	#Used by the progress calculator - this is the number of rows in the table
		foreach ($value in $values)
		{
			for ($i = 0; $i -le 3; $i++)
			{
				#Cater for skipping the spacer column down the centre & bold the headers
				switch ($i)
				{
					'0'
					{
						$column = $i + 1
						if (!$wdTemplate) { $table.cell($row,$column).Shading.BackgroundPatternColor = $iLabelColour } # $wdColorFirstRow
						$table.cell($row,$column).Range.Font.Bold = $false
					}
					'1'
					{
						$column = $i + 1
					}
					'2'
					{
						$column = $i + 2
						if (!$wdTemplate) { $table.cell($row,$column).Shading.BackgroundPatternColor = $iLabelColour } # $wdColorFirstRow
						$table.cell($row,$column).Range.Font.Bold = $false
					}
					'3'
					{
						$column = $i + 2
					}
				}
				if (($value[$i] -eq $null) -or ($value[$i] -eq '')) { continue }
				$table.cell($row,$column).Range.text = ($value[$i]).ToString() #ToString added as Win10 would crash here unable to cast System.Int32 to System.String.
				if ($progressBarTitle -ne '')
				{
					#We only update the progress bar if we have a value in this cell. (Speeds execution?)
					$progressPercent2 = ((($row - 1) * 4) + ($i+1)) / ($cellcount * 4) * 100
					If ($progressPercent2 -gt 100) { $progressPercent2 = 100 } #Trap any accidental percentage overrun
					write-progress -id 2 -parentid 1 -Activity 'Table' -Status $($progressBarTitle) -PercentComplete $progressPercent2
				}
			}
			#OK, the row's complete - do we need to consolidate any columns into titles/headers?
			if ($value[0] -eq 'SPAN')
			{
				$table.cell($row,1).Range.text = ''	#Strip the SPAN code
				$table.cell($row,1).Merge($table.cell($row,5))
				$table.cell($row,1).Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
				if (!$wdTemplate) { $table.cell($row,1).Shading.BackgroundPatternColor = $iTitleColour }
				$table.cell($row,1).Range.Font.Bold = $true
			}
			if ($value[2] -eq 'SPAN-R')
			{
				$table.cell($row,4).Range.text = '' #Strip the SPAN code
				$table.cell($row,4).Merge($table.cell($row,5))
				$table.cell($row,4).Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
				if (!$wdTemplate) { $table.cell($row,4).Shading.BackgroundPatternColor = $iTitleColour }
				$table.cell($row,4).Range.Font.Bold = $true
			}
			if ($value[0] -eq 'SPAN-L')
			{
				$table.cell($row,1).Range.text = ''	#Strip the SPAN code
				$table.cell($row,1).Merge($table.cell($row,2))
				$table.cell($row,1).Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
				if (!$wdTemplate) { $table.cell($row,1).Shading.BackgroundPatternColor = $iTitleColour }
				$table.cell($row,1).Range.Font.Bold = $true
			}
			$row ++
		}
		if ($progressBarTitle -ne '')
		{
			write-progress -id 2 -parentid 1 -Activity 'Table' -Status $($progressBarTitle) -PercentComplete (100) -completed
		}
		#Step out of the table:
		$null = $Selection.EndKey(6, 0)
		$Selection.TypeParagraph() #Stick a blank line after each table

	}


	function Write-AtBookmark()
	{
		param ([string]$bookmark, [string]$data)

		if ($doc.Bookmarks.Exists($bookmark))
		{
			$Range = $doc.Bookmarks.Item($bookmark).Range
			$Range.Text = $data
			$null = $doc.Bookmarks.Add($bookmark,$Range)
		}
		else
		{
			write-warning -Message ('Bookmark ""{0}"" not found' -f $bookmark)
		}
	 }


	function Decode-Sequence ()
	{
		param ([string]$EncodedSequence , [int]$byteCount)
		[byte[]] $SequenceArray = [Convert]::FromBase64CharArray($EncodedSequence,0,$EncodedSequence.Length)
		[Int16[]] $DecodedSequence = @()

		#Depending on the sequence type, we either return now or need to decode it further
		if ($byteCount -eq 1)
		{
			return $SequenceArray
		}
		else
		{
			#Convert the members of the dual byte array back to a UInt16
			for ($i=0; $i -le $SequenceArray.Count; $i=$i+2)
			{
				$j = $i + 1
				[Int16] $SequenceMember = ($SequenceArray[$j] * 256)
				$SequenceMember = $SequenceMember + ($SequenceArray[$i])
				if ($SequenceMember -ne 0) { $DecodedSequence += $SequenceMember} #Skip any value that references entry 0
			}
			return ,$DecodedSequence
		}
	}


	function Fix-NullDescription ()
	{
		param ([string]$TableDescription , [string]$TableValue, [string]$TablePrefix)
		if (($TableDescription -eq $null) -or ($TableDescription -eq ''))
		{
			return ($TablePrefix + $TableValue)
		}
		else
		{
			return ($TableDescription)
		}
	}


	function Strip-TrailingCR ()
	{
		param ([string]$DelimitedString)
		while (1)
		{
			if ($DelimitedString.Length -ge 1)
			{
				if ($DelimitedString.Substring($DelimitedString.Length -1) -eq "`n")
				{
					$DelimitedString = $DelimitedString.Substring(0,$DelimitedString.Length-1)
				}
				else
				{
					break
				}
			}
			else
			{
				break
			}
		}
		return $DelimitedString
	}


	function Consolidate-ChannelList ()
	{
		param ([string]$Channels)
		$ChArray = ($Channels).Split(',')
		$card, [int]$lastport = ($ChArray[0]).Split(':')	#Suck out the first entry as a starting position
		$ChArray = $ChArray[1..$ChArray.Length]				#Now remove it from the array

		$ConsolidatedList = ('{0}:{1}' -f ($card), ($lastport))	#Initialise the string with the first port
		$RunningSequence = $false

		foreach ($channel in $ChArray)
		{
			$disCard, [int]$port = ($channel).Split(':')	#We don't use the card here, just the port - hence 'disCard'. ;-)
			if ($port -eq ($lastport + 1))
			{
				#is port sequential?
				$lastport = $port	#Update LastPort
				$RunningSequence = $true
			}
			else
			{
				if ($RunningSequence -eq $true)
				{
					$ConsolidatedList += (' - {0}:{1}, {0}:{2}' -f ($card), ($lastport), ($port))
					$RunningSequence = $false
				}
				else
				{
					$ConsolidatedList += (', {0}:{1}' -f ($card), ($port))
				}
				$lastport = $port	#Update LastPort
			}
		}
		#If we have a sequence, the last entry is still in memory:
		if ($RunningSequence -eq $true)
		{
			$ConsolidatedList += (' - {0}:{1}' -f ($card), ($lastport))
		}
		return $ConsolidatedList
	}


	# Returns '<n/a this rls>' if the passed value is $null, otherwise returns the value, or the result of a lookup (if $LookupTable is populated)
	# Aha: 3/5/16. Adding " [AllowEmptyString()] in the parameter declarations might get me around that. I'll save this for later.
	function Test-ForNull ()
	{
		param ([hashtable]$LookupTable, $value) # If I cast $value as a string, a null will come through here as a '', which I don't want.
												# By leaving it 'un-cast' we can accept a null & still also correctly receive & handle an empty string. (Thank you StackOverflow!)

		if ($value -eq $null)
		{
			return '<n/a this rls>'
		}
		else
		{
			if (($LookupTable.Count -eq 0) -or ($LookupTable -eq $null)) # If we passed $null for the LookupTable, its count will be 0
			{
				return $value
			}
			else
			{
				$LookupResult = $LookupTable.Get_Item($value)
				if ($LookupResult -eq "")
				{
					return "ScriptErr:$(value)" #If the lookup failed to return a value, add a marker in the doc
				}
				else
				{
					return $LookupResult
				}
			}
		}
	}


	function Extract-FromArchive
	{
		param ([string] $Action, [string] $InputFile, [string] $FileToExtract, [string] $ReturnType)

		if ($action -eq 'Extract')
		{
			[String]$arguments = ('e -ttar "{0}" -aoa "{1}" -y -r -so' -f $Inputfile, $FileToExtract)
		}
		else
		{
			[String]$arguments = ('l -ttar "{0}" "{1}" -r' -f $Inputfile, $FileToExtract)
		}

		# http://blog.danskingdom.com/powershell-function-to-create-a-password-protected-zip-file/
		# Look for the 7zip executable:
		$pathTo32Bit7Zip = 'C:\Program Files (x86)\7-Zip\7z.exe'
		$pathTo64Bit7Zip = 'C:\Program Files\7-Zip\7z.exe'
		$pathToStandAloneExe = Join-Path -Path $dir -ChildPath '7za.exe' # (We set $dir earlier to be the script's directory)
		if 	   ( Test-Path -Path $pathTo64Bit7Zip	 ) { $pathTo7ZipExe = $pathTo64Bit7Zip }
		elseif ( Test-Path -Path $pathTo32Bit7Zip	 ) { $pathTo7ZipExe = $pathTo32Bit7Zip }
		elseif ( Test-Path -Path $pathToStandAloneExe ) { $pathTo7ZipExe = $pathToStandAloneExe }
		else   { throw 'Could not find the 7-zip executable. Get it from http://www.7-zip.org/.' }

		# The object here (and hence the complexity) is to extract 'symphonyconfig.xml' directly into memory
		# http://stackoverflow.com/questions/8761888/powershell-capturing-standard-out-and-error-with-start-process
		$pinfo = New-Object -TypeName System.Diagnostics.ProcessStartInfo
		$pinfo.FileName = $pathTo7ZipExe
		$pinfo.RedirectStandardError = $true
		$pinfo.RedirectStandardOutput = $true
		$pinfo.UseShellExecute = $false
		$pinfo.Arguments = $arguments
		$p = New-Object -TypeName System.Diagnostics.Process
		$p.StartInfo = $pinfo
		$null = $p.Start()
		switch ($ReturnType)
		{
			'xml'   { $returnValue = [xml] $p.StandardOutput.ReadToEnd()	}
			default { $returnValue = [object] $p.StandardOutput.ReadToEnd() }
		}

		#If the file was not unzipped successfully:
		if (!(($p.HasExited -eq $true) -and ($p.ExitCode -eq 0)))
		{
			throw ("There was a problem extracting '{0}' from the archive '{1}'." -f $FileToExtract, [IO.Path]::GetFileName($InputFile))
		}
		else
		{
			If ($returnValue.length -eq 0)
			{
				write-warning -Message ('{0} not found in .tar' -f ($FileToExtract))
			}
			else
			{
				write-verbose -message ('{0} extracted from .tar OK' -f ($FileToExtract))
			}
		}
		return $returnValue
	}


	function Convert-RgbToWordColour ()
	{
		param ([string] $RGBString, [int32] $DefaultColour)
		if ($rgbstring -eq '')
		{
			return $defaultcolour
		}
		else
		{
			[Int32[]]$rgbarray = ($rgbstring).split(',')
			# Thank you Sophie for this code & the suggestion:
			# transformation rgb -> wdcolor
			# wdcolor = (red + 0x100 * green + 0x10000 * blue)
			# light orange (255,204,153)
			try
			{
				[Int32]$rgbcolour = $rgbarray[0] + (0x100 * $rgbarray[1]) + (0x10000 * $rgbarray[2])
				if ($rgbcolour -gt 16777215) { write-warning -message 'Invalid colour value passed - default used instead'; return $defaultcolour}
				return $rgbcolour
			}
			catch
			{
				write-warning -message 'Invalid colour value passed - default used instead'
				return $defaultcolour
			}
		}
	}


	function Fix-NullAndDuplicateDescriptions()
	{
		param ($XmlObjectList, [string]$TableListClassName, [string]$TablePrefix)
		$ObjectLookup = @{}
		# First we build a hash table of Indexes and Descriptions. This also fixes null descriptions:
		foreach ($XmlObject in $XmlObjectList)
		{
			if ($XmlObject.IE.classname -eq $TableListClassName)
			{
				if ($XmlObject.IE.Description -ne '')
				{
					$ObjectLookup.Add($XmlObject.value, $XmlObject.IE.Description)
				}
				else
				{
					$ObjectLookup.Add($XmlObject.value, $TablePrefix + $XmlObject.value)
				}
			}
		}
		# Now de-dupe (Thank you SO: http://stackoverflow.com/questions/24454626/how-to-find-duplicate-values-in-powershell-hash)
		$dupes = $ObjectLookup.GetEnumerator() | Group-Object -Property Value | Where-Object { $_.Count -gt 1 }
		foreach ($dupe in $dupes.Group)
		{
			if ($dupe.Name -ne $null)
			{
				$ObjectLookup.Set_Item($dupe.Name,($ObjectLookup.Get_Item($dupe.Name) + ' (' + $dupe.Name + ')'))
			}
		}
		return $ObjectLookup
	}


	function Convert-MaskToCIDR()
	{
		# https://www.experts-exchange.com/questions/27258790/Calculating-CIDR-notation-with-PowerShell.html
		param ([string]$Mask)
		$MaskArray = $Mask.split('.')
		$CIDR = [int] 0
		$Octet = [int]0
		Foreach ($Octet in $MaskArray)
		{
			if ($Octet -eq 255){$CIDR += 8}
			if ($Octet -eq 254){$CIDR += 7}
			if ($Octet -eq 252){$CIDR += 6}
			if ($Octet -eq 248){$CIDR += 5}
			if ($Octet -eq 240){$CIDR += 4}
			if ($Octet -eq 224){$CIDR += 3}
			if ($Octet -eq 192){$CIDR += 2}
			if ($Octet -eq 128){$CIDR += 1}
		}
		return $CIDR
	}


	function Dump-ErrToScreen ()
	{
		if ($Global:Debug)
		{
			$Global:error[0] | Format-List -Property * -Force #This dumps to screen as white for the time being. I haven't been able to get it to dump in red
		}
		else
		{
			write-host ('{0}' -f ($Global:error[0])) -ForegroundColor Red -BackgroundColor Black
		}
	}


	function Get-UpdateInfo
	{
	  <#
		  .SYNOPSIS
		  Queries an online XML source for version information to determine if a new version of the script is available.
		  *** This version customised by Greig Sheridan. @greiginsydney https://greiginsydney.com ***

		  .DESCRIPTION
		  Queries an online XML source for version information to determine if a new version of the script is available.

		  .NOTES
		  Version			   : 1.2 - See changelog at https://ucunleashed.com/3168 for fixes & changes introduced with each version
		  Wish list			 : Better error trapping
		  Rights Required	   : N/A
		  Sched Task Required   : No
		  Lync/Skype4B Version  : N/A
		  Author/Copyright	  : © Pat Richard, Office Servers and Services (Skype for Business) MVP - All Rights Reserved
		  Email/Blog/Twitter	: pat@innervation.com  https://ucunleashed.com  @patrichard
		  Donations			 : https://www.paypal.me/PatRichard
		  Dedicated Post		: https://ucunleashed.com/3168
		  Disclaimer			: You running this script/function means you will not blame the author(s) if this breaks your stuff. This script/function
								is provided AS IS without warranty of any kind. Author(s) disclaim all implied warranties including, without limitation,
								any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use
								or performance of the sample scripts and documentation remains with you. In no event shall author(s) be held liable for
								any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss
								of business information, or other pecuniary loss) arising out of the use of or inability to use the script or
								documentation. Neither this script/function, nor any part of it other than those parts that are explicitly copied from
								others, may be republished without author(s) express written permission. Author(s) retain the right to alter this
								disclaimer at any time. For the most up to date version of the disclaimer, see https://ucunleashed.com/code-disclaimer.
		  Acknowledgements	  : Reading XML files
								http://stackoverflow.com/questions/18509358/how-to-read-xml-in-powershell
								http://stackoverflow.com/questions/20433932/determine-xml-node-exists
		  Assumptions		   : ExecutionPolicy of AllSigned (recommended), RemoteSigned, or Unrestricted (not recommended)
		  Limitations		   :
		  Known issues		  :

		  .EXAMPLE
		  Get-UpdateInfo -Title 'Convert-RibbonSbcConfigToWord.ps1'

		  Description
		  -----------
		  Runs function to check for updates to script called 'Convert-RibbonSbcConfigToWord.ps1'.

		  .INPUTS
		  None. You cannot pipe objects to this script.
	  #>
		[CmdletBinding(SupportsShouldProcess = $true)]
		param (
		[string] $title
		)
		try
		{
			[bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
			if ($HasInternetAccess)
			{
				write-verbose -message 'Performing update check'
				# ------------------ TLS 1.2 fixup from https://github.com/chocolatey/choco/wiki/Installation#installing-with-restricted-tls
				$securityProtocolSettingsOriginal = [Net.ServicePointManager]::SecurityProtocol
				try {
				  # Set TLS 1.2 (3072). Use integers because the enumeration values for TLS 1.2 won't exist in .NET 4.0, even though they are
				  # addressable if .NET 4.5+ is installed (.NET 4.5 is an in-place upgrade).
				  [Net.ServicePointManager]::SecurityProtocol = 3072
				} catch {
				  write-verbose -message 'Unable to set PowerShell to use TLS 1.2 due to old .NET Framework installed.'
				}
				# ------------------ end TLS 1.2 fixup
				[xml] $xml = (New-Object -TypeName System.Net.WebClient).DownloadString('https://greiginsydney.com/wp-content/version.xml')
				[Net.ServicePointManager]::SecurityProtocol = $securityProtocolSettingsOriginal #Reinstate original SecurityProtocol settings
				$article  = select-XML -xml $xml -xpath ("//article[@title='{0}']" -f ($title))
				[string] $Ga = $article.node.version.trim()
				if ($article.node.changeLog)
				{
					[string] $changelog = 'This version includes: ' + $article.node.changeLog.trim() + "`n`n"
				}
				if ($Ga -gt $ScriptVersion)
				{
					$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
					$updatePrompt = $wshell.Popup(("Version {0} is available.`n`n{1}Would you like to download it?" -f ($ga), ($changelog)),0,'New version available',68)
					if ($updatePrompt -eq 6)
					{
						Start-Process -FilePath $article.node.downloadUrl
						write-warning -message "Script is exiting. Please run the new version of the script after you've downloaded it."
						exit
					}
					else
					{
						write-verbose -message ('Upgrade to version {0} was declined' -f ($ga))
					}
				}
				elseif ($Ga -eq $ScriptVersion)
				{
					write-verbose -message ('Script version {0} is the latest released version' -f ($Scriptversion))
				}
				else
				{
					write-verbose -message ('Script version {0} is newer than the latest released version {1}' -f ($Scriptversion), ($ga))
				}
			}
			else
			{
			}

		} # end function Get-UpdateInfo
		catch
		{
			write-verbose -message 'Caught error in Get-UpdateInfo'
			if ($Global:Debug)
			{
				$Global:error | Format-List -Property * -Force #This dumps to screen as white for the time being. I haven't been able to get it to dump in red
			}
		}
	}


	function Sequence-SmmElements
	{
		param (
		$DescList,
		$UriUList,
		$UriPList,
		$HeaderList,
		$HeaderParamList,
		[string] $RuleType
		)

		$Desclist += $HeaderList #These can be joined together - the sequencing will pull them out in the correct order, and they're never both present on the same element/rule

		$ReturnValue = ''

		#The challenge here is that the SMM 'element classes' - the parameter changes that make up each of the rules - aren't in
		# any particular order in the XML, so I need to rearrange them into the sequence in which they appear on-screen.

		#For bonus points, there's nothing in the XML to tell what each of the elements "<ElementClass>9</ElementClass>" is for - its all in the context
		# That's required some of the extra complexity here.

		$Sequence = @()
		$SequenceOrder = ('<Default>', 'Method:', 'Display Name:', 'URI:', 'URI Scheme:', 'URI USER INFO:', 'URI User:', 'Password:', 'URI User Parameters:', `
							'URI Host:', 'URI Port:', 'URI Parameters:', 'Header Parameters:', 'SIP Version:')

		foreach ($SequenceItem in $SequenceOrder)
		{
			switch ($SequenceItem)
			{
				'URI User Parameters:'
				{
					foreach ($Element in $UriUList)
					{
						$value = [regex]::replace($SIPMessageRuleElementLookup.Get_Item($Element), 'URI/Header Parameters' , 'URI User Parameters')
						$returnValue += $value + "`n"
					}
				}

				'URI Parameters:'
				{
					foreach ($Element in $UriPList)
					{
						$value = [regex]::replace($SIPMessageRuleElementLookup.Get_Item($Element), 'URI/Header Parameters' , 'URI Parameters')
						$returnValue += $value + "`n"
					}
				}

				default
				{
					foreach ($Element in $DescList)
					{
						$value = $SIPMessageRuleElementLookup.Get_Item($Element) + "`n"
						if ($value.StartsWith($SequenceItem))
						{
							$value = [regex]::replace($SIPMessageRuleElementLookup.Get_Item($Element), '<Default>' , $RuleType)
							$returnValue += $value + "`n"
						}
					}
				}
			}
		}

		#Header parameters, if present, are always the last to show:
		foreach ($Element in $HeaderParamList)
		{
			$value = [regex]::replace($SIPMessageRuleElementLookup.Get_Item($Element), 'URI/Header Parameters' , 'Header Parameters')
			$returnValue += $value + "`n"
		}

		return $ReturnValue
	}


	function Do-Convert
	{
		param (
		[string] $InputFile,
		[string] $OutputFile,
		[switch] $IncludeNodeInfo,
		[string] $NodeID,
		[string] $HardwareID,
		[switch] $MakePDF,
		[switch] $SkipWrite,
		[String[]] $Do='All',
		[switch] $Landscape,
		[switch] $RedactIP,
		[string] $TitleColour,
		[string] $LabelColour,
		[string] $WdTemplate,
		[string] $WdTableStyle,
		[switch] $SonusGreen
		)

		$xml = ''
		$nodeInfo = ''
		$MOHFilename = '(Not Available)'	#We'll populate this if the user provides a .tar and there's a MOH file in there!

		if (!$InputFile)
		{
			write-verbose -message 'InputFile  not specified. "symphonyconfig.xml"  assumed'
			$InputFile = 'symphonyconfig.xml'
		}

		if ([IO.Path]::IsPathRooted($InputFile))
		{
			#It's absolute. Safe to leave.
		}
		else
		{
			#It's relative.
			$InputFile = [IO.Path]::GetFullPath((Join-Path -path $dir -childpath $InputFile))
		}

		if (!$OutputFile)
		{
			$Outputfile = ([IO.Path]::ChangeExtension($InputFile, 'docx'))
			write-verbose -message ('OutputFile not specified. "{0}" assumed' -f ($OutputFile))
		}

		if ([IO.Path]::IsPathRooted($Outputfile))
		{
			#It's absolute. Safe to leave.
		}
		else
		{
			#It's relative.
			$Outputfile = [IO.Path]::GetFullPath((Join-Path -path $dir -childpath $Outputfile))
		}

		$RawOutputFileName = [IO.Path]::GetFileName($OutputFile)

		if (!$NodeId)
		{
			$NodeId = '<Not captured>'
		}
		if (!$HardwareID)
		{
			$HardwareID = '<Not captured>'
		}

		#Is 'inputfile' an XML or an archive?
		if ([IO.Path]::GetExtension($InputFile) -eq '.tar')
		{
			write-verbose -message 'InputFile is .tar. Attempting to extract file from archive'

			$xml = Extract-FromArchive -Action 'Extract' -InputFile $InputFile -FileToExtract 'symphonyconfig.xml' -ReturnType 'xml'
			$nodeInfo = (Extract-FromArchive -Action 'Extract' -InputFile $InputFile -FileToExtract 'nodeinfo.txt' -ReturnType 'object')

			#First we get a list of the WAVs in the archive, then strip everything to leave the remaining MOH filename
			$MOHFilename = Extract-FromArchive -Action 'List' -InputFile $InputFile -FileToExtract '*.wav' -ReturnType 'object'
			$MOHFilename = [regex]::match($MOHFilename, '\.\\moh\\(.*.wav)').Groups[1].Value
			if ($MOHFilename -eq '') { $MOHFilename = '<No file found>' }
		}
		else
		{
			write-verbose -message "InputFile is not .tar. Assuming it's an xml file"
			if ($IncludenodeInfo)
			{
				$SourcePath = split-path -Path $InputFile
				$NodeInfoPath = Join-Path -path (split-path -Path $InputFile) -childpath 'NodeInfo.txt'
				if (test-path -Path $NodeInfoPath)
				{
					#$nodeInfo = get-content $NodeInfoPath -raw 			 #P$ v3 version
					$nodeInfo = [IO.File]::ReadAllText($NodeInfoPath) #Safe for P$ v2.
				}
				else
				{
					#Nope, no NodeInfo file in this dir
					write-warning -message ('NodeInfo file not found at ""{0}\""' -f ($SourcePath))
					$nodeInfo = ''
				}
			}
		}

		$DoAll = $false
		$DoCalls = $false
		$DoSig = $false
		$DoIP = $false
		$DoSIP = $false
		$DoSystem = $false
		$DoMisc = $false
		$DoMaint = $false
		$DoList = ''

		if ($Do)
		{
			foreach ($Item in $Do)
			{
				switch ($item)
				{
					All	{$DoAll = $true}
					Calls  {$DoCalls = $true}
					Sig	{$DoSig = $true}
					IP	 {$DoIP = $true}
					SIP	{$DoSIP = $true}
					Sys	{$DoSystem = $true}
					System {$DoSystem = $true}
					Misc   {$DoMisc = $true}
					Maint  {$DoMaint = $true}
				}
				$DoList += $item + ', '
			}
			$DoList = $DoList.Substring(0,$DoList.Length-2) #Strip the trailing comma (& invisible space)
			$DoList = $DoList.ToLower()
			if ($Do -contains 'All')
			{
				$DoList = ''	#Null the list to ensure we don't write 'incomplete: only contains: All' after the TOC
			}
			else
			{
				write-verbose -message ('Only writing subset: {0}' -f ($DoList))
			}
		}

		#--------------------------------
		# open the backup file ----------
		#--------------------------------
		write-verbose -message 'Opening the backup file'
		try
		{
			if ($xml -eq '')
			{
				$xml = [xml] (get-content -Path $InputFile -Encoding UTF8 -ErrorAction stop)
			}
		}
		catch
		{
			write-error -Message ("Item not found or file is corrupt:`n`n{0}" -f $_)
			return
		}

		#---------------------------------
		# Open Word & create new document
		#---------------------------------
		write-verbose -message 'Opening Word'
		write-progress -id 1 -Activity 'Initialising' -Status 'Opening Word' -PercentComplete (2)

		if ($SonusGreen)
		{
			# The below will override the default 'Sonus green' colour scheme if the user has provided alternatives:
			$iTitleColour = Convert-RgbToWordColour -RGBString $TitleColour -DefaultColour $wdColorLightGreen
			$iLabelColour = Convert-RgbToWordColour -RGBString $LabelColour -DefaultColour $beLightGreen
		}
		else
		{
			# The below will override the default 'Ribbon grey' colour scheme if the user has provided alternatives:
			$iTitleColour = Convert-RgbToWordColour -RGBString $TitleColour -DefaultColour 13816530 # #D2D2D2 (Alt #FCFCFC = 11711154 )
			$iLabelColour = Convert-RgbToWordColour -RGBString $LabelColour -DefaultColour 16579836 # #B2B2B2
		}

		#The bulk of the script exists inside a try/finally loop to ensure Word is closed if the user aborts
		try
		{
			try
			{
				$word = new-object -ComObject Word.Application  -verbose:$false	# Removed quotes around Word.Application 1/6/15 chasing [ref] error
				[ref]$saveFormat = 'microsoft.office.interop.word.WdSaveFormat' -as [type]
			}
			catch [Runtime.InteropServices.COMException]
			{
				write-host "Looks like Word isn't installed. Aborting with fatal error:" -ForegroundColor Red -BackgroundColor Black
				Dump-ErrToScreen
				break #We can safely exit now - no files are open and no Word was able to be launched
			}
			catch
			{
				write-host 'Error opening Word. Aborting with fatal error:' -ForegroundColor Red -BackgroundColor Black
				Dump-ErrToScreen
				break #We can safely exit now - no files are open and no Word was able to be launched
			}
			# Reference: https://msdn.microsoft.com/en-us/library/bb238158%28v=office.12%29.aspx
			if ($word -eq $null) { write-warning -message "`$word is Null - this isn't going to end well!" } # We should never reach this line!
			try
			{
				if ($wdTemplate)
				{
					if (test-Path -Path $wdTemplate)
					{
						$doc = $word.Documents.Add($wdTemplate)
					}
					else
					{
						write-warning -message ('Word template file ""{0}"" not found. Using the Default template' -f ($wdTemplate))
						$wdTemplate = $null
						$doc = $word.Documents.Add()
					}
				}
				else
				{
					$doc = $word.Documents.Add()
				}
			}
			catch
			{
				write-host 'Error creating new Word document. Fatal error:' -ForegroundColor Red -BackgroundColor Black
				Dump-ErrToScreen
				throw
			}
			#word.Visible=$True #Toggle these to watch all the excitement on-screen
			$word.Visible=$False
			try
			{
				$doc.ShowSpellingErrors = $false
				$doc.ShowGrammaticalErrors = $false
			}
			catch
			{
				write-warning "Word threw an error disabling spelling/grammar checking"
			}
			try
			{
				$doc.SaveAs([ref] $OutputFile, [ref]$saveFormat::wdFormatDocument) # https://blogs.technet.microsoft.com/heyscriptingguy/2008/11/11/hey-scripting-guy-how-can-i-create-a-microsoft-word-document-from-wmi-information/
			}
			catch [Management.AutomationInvocationInfo]
			{
				if ($_.Exception -match 'Word cannot complete the save due to a file permission error')
				{
					write-host ('File permission error saving Word document ""{0}"".' -f ($OutputFile)) -ForegroundColor Red -BackgroundColor Black
					write-host "Check the file isn't open in another instance of Word - check Task Manager. Fatal error:" -ForegroundColor Red -BackgroundColor Black
					Dump-ErrToScreen
					throw
				}
			}
			catch
			{
				# OK, this is an ugly kludge, but if it doesn't want to Save using '[ref]', let's try again without it:
				try
				{
					$doc.SaveAs($OutputFile, $saveFormat::wdFormatDocument)
				}
				catch
				{
					write-host ('Error saving new Word document ""{0}"". Fatal error:' -f ($OutputFile)) -ForegroundColor Red -BackgroundColor Black
					Dump-ErrToScreen
					throw
				}
			}

			if ($WdTableStyle)
			{
				$tableExists = 0
				foreach ($style in $doc.styles)
				{
					if ($style.NameLocal -contains $WdTableStyle)
					{
						if ($style.Table -ne $null)
						{
							#Yes! It's a valid Table style
							$tableExists = 1
							break
						}
					}
				}
				if (!$tableExists)
				{
					write-warning -message ('Table style ""{0}"" not found in the document template' -f $WdTableStyle)
					$WdTableStyle = $null
				}
			}

			#--------------------------------
			# Build the title page ----------
			#--------------------------------

			write-verbose -message 'Creating the title page & TOC'
			write-progress -id 1 -Activity 'Initialising' -Status 'Creating the title page & TOC' -PercentComplete (3)

			$selection=$word.Selection
			if (!$wdTemplate)
			{
				$selection.Font.Name = 'Arial'
				$selection.Font.Color = $wdColorAutomatic
			}
			$selection.Font.Size = 24
			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$selection.TypeText('SBC As-Built')
			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$selection.Font.Size=16
			$selection.TypeText('Description')
			$selection.TypeText("`t")
			$null = $doc.Bookmarks.Add('sysDescription', $selection)
			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$selection.TypeText('Location')
			$selection.TypeText("`t`t")
			$null = $doc.Bookmarks.Add('sysLocation', $selection)
			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$selection.TypeText('Contact')
			$selection.TypeText("`t`t")
			$null = $doc.Bookmarks.Add('sysContact', $selection)
			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$selection.Font.Size=14
			$selection.TypeText('Release')
			$selection.TypeText("`t`t")
			$null = $doc.Bookmarks.Add('sysRelease', $selection)
			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$selection.TypeText('Platform')
			$selection.TypeText("`t`t")
			$null = $doc.Bookmarks.Add('sysPlatform', $selection)
			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$selection.TypeText('Created from')
			$selection.TypeText("`t")
			$selection.TypeText([IO.Path]::GetFileName($InputFile))
			$selection.TypeParagraph()
			$selection.TypeText('on ')
			$selection.TypeText("`t`t`t")
			$selection.TypeText((Get-Date -Format 'd MMM yyyy').ToString())
			$selection.InsertNewPage()

			#--------------------------------
			# Table Of Contents -------------
			#--------------------------------
			if (!$wdTemplate)
			{
				$selection.Font.Size=20
				$selection.Font.Name='Arial'
				$selection.Font.Color = $wdColorAutomatic
			}
			$selection.TypeText('Table of Contents')
			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$rngTOC = $selection.Range

			$null = $doc.TablesOfContents.Add($rngTOC, $wdMissingValue, $wdMissingValue, $wdMissingValue, $wdMissingValue, $wdMissingValue, $wdMissingValue, $wdMissingValue, $wdMissingValue, 1, $wdMissingValue, $wdMissingValue)

			$selection.TypeParagraph()
			$selection.TypeParagraph()
			$selection.Style= $wdStyleNormal
			$selection.TypeText('Created by a PowerShell script from https://greiginsydney.com/uxbuilder')
			$selection.TypeParagraph()
			$selection.TypeText(("Version`t{0}" -f $ScriptVersion))
			if ($DoList -ne '')
			{
				$selection.TypeParagraph()
				$selection.TypeParagraph()
				$selection.TypeText(('Captured information is incomplete. Only contains: {0}' -f $DoList))
				$selection.TypeParagraph()
			}

			if ($Landscape)
			{
				 $selection.InsertBreak($wdPageBreak)
			}
			else
			{
				$selection.InsertBreak($wdSectionBreakNextPage)
			}
			$selection.PageSetup.Orientation = $wdOrientLandscape

			#--------------------------------------------
			# READING FROM THE BACKUP FILE(S) STARTS HERE --
			#--------------------------------------------
			# Current versions of the firmware have a 'SystemRelease' node that captures the operating firmware version.
			# It's not present in very old (v1?) firmware, so we set a default value if it doesn't exist.

			$config = $xml.selectsinglenode('/Configuration/SystemRelease')
			[string]$release = $config.'#text'

			$release = [regex]::replace($release, 'v' , ' v') # Add a space between version & build
			if (($release -eq '') -or ($release -eq $null))	#Early versions (v1.x) don't capture the release
			{
				write-verbose -message 'System Release  = <Not Available>'
				$release = '<Not Available - too old>'
				$releaseBuild = 0
				$releaseVersion = 0
			}
			else
			{
				$releaseVersion = [int](($release).Split('.v'))[0] # Turns '6.1.0 v457' into just 6
				$releaseBuild   = [int](($release).Split('v'))[1] # Turns '6.1.0 v457' into just 457
				write-verbose -message ('System Release  = {0}' -f ($release))
			}
			Write-AtBookmark -bookmark 'SysRelease' -data $release

			$config = $xml.selectsinglenode('/Configuration/PlatformType')
			[string]$platform = $config.'#text'

			$platform = [regex]::replace($platform, 'SBC' , 'SBC ') # Add a space between 'SBC' and model
			$platform = [regex]::replace($platform, 'SWeLite' , 'SWe Lite') # Add a space between 'SBC' and model

			if ($platform -eq 'SWe Lite') { $SWeLite = 1 } else { $SWeLite = 0 } # I test for SWeLite SO much it's easier to refer to a boolean

			if (($platform -eq '') -or ($platform -eq $null))	#This was only added with v4.0
			{
				write-verbose -message 'System platform = <Not Available>'
				$platform = '<Not Available - too old>'
			}
			else
			{
				write-verbose -message ('System platform = {0}' -f ($platform))
			}
			Write-AtBookmark -bookmark 'Sysplatform' -data $platform

			#---------------------------------------------------------------------------------------------------------
			# If available, read the NodeInfo.txt file and capture its contents. Various bits need to be stuffed into
			# tables in the sections below.
			#---------------------------------------------------------------------------------------------------------

			$PortLicenceCollection = @()
			$PortLicenceColumnTitles = @('Feature', 'Licensed', 'Number of Licensed Ports')

			$FeatureLicenceCollection = @()
			
			#---------------------------------------------------------------------------------------------------------
			# The Lite has five columns in this table but the 1k/2k only four. As of v8.1.5 I'm treating all values as
			# though they have five columns, then leaving it to the very end to decide how many to write to file.
			# This makes for cleaner code and less risk of errors.
			#---------------------------------------------------------------------------------------------------------
			if ($SweLite)
			{
				$FeatureLicenceColumnTitles = @('Feature', 'Licensed', 'Total Licenses', 'Available Licenses', 'Feature Expiration')
			}
			else
			{
				$FeatureLicenceColumnTitles = @('Feature', 'Licensed', 'Total Licenses', 'Available Licenses')
			}
			$LicenceExpiration = ''
			$InventoryOf = ''
			$WindowsFactoryLicence = '<Not Available>'
			$SoftwareBundled = '<Not Available>'
			$NodeLicenseSKU = '<Not Available>'

			$ASM_WindowsEthernetMainName = 'Not Available' # Not actually used anywhere yet
			$ASM_WindowsEthernetMainMac  = 'Not Available' # Not actually used anywhere yet
			$ASM_WindowsEthernetSecName  = 'Not Available' # Not actually used anywhere yet
			$ASM_WindowsEthernetSecMac   = 'Not Available' # Used in ASM section
			$ASM_version = 'Not Available'

			If ($nodeInfo -ne '')
			{
				if ($IncludeNodeInfo)
				{
					write-verbose -message 'Reading NodeInfo.txt'
				}
				else
				{
					write-verbose -message '.TAR provided - reading NodeInfo.txt'
				}

				#$NodeInfo.GetType()
				$NodeInfoArray = ($NodeInfo).ToString().Split([environment]::NewLine)
				ForEach ($nodeLine in $NodeInfoArray)
				{
					$TempNode = (($nodeLine).Split(':'))
					$TempNode0 = $TempNode[0].Trim()
					if (($TempNode.count -eq 1) -and ($TempNode0 -notmatch 'License Information')) #There's nothing (else) of interest on this line. Read the next one.
					{
						continue
					}
					if ($TempNode.count -ge 2)
					{
						$TempNode1 = $TempNode[1].Trim()
					}
					else
					{
						$TempNode1 = '' #Should be redundant.
					}
					switch ($TempNode0)
					{
						{(($_ -eq 'Node Serial Number') -or ($_ -eq 'SWe Lite ID'))}
						{
							if ($NodeId -ne '<Not captured>')
							{
								#Then the user has provided one.
								If ($NodeId -ne $TempNode1)
								{
									write-warning -message ("User-provided NodeID of '{0}' differs from value of '{1}' from file. User value discarded." -f ($NodeID), ($TempNode1))
								}
							}
							$NodeId = $TempNode1
						}

						#We'll check the Hostname from NodeInfo in the initial parse and abort if it differs from what's in the XML - we assume user has mixed up XML & NodeInfo.
						'Hostname'
						{
							$NodeHostname = $TempNode1
						}

						'Hardware ID'
						{
							if ($HardwareID -ne '<Not captured>')
							{
								#Then the user has provided one.
								If ($HardwareID -ne $TempNode1)
								{
									write-warning -message ("User-provided HardwareID of '{0}' differs from value of '{1}' from file. User value discarded." -f ($HardwareID), ($TempNode1))
								}
							}
							$HardwareID = $TempNode1
						}

						#All of the inventory sections start with a header. This string is updated with each header so we know what item the values are referring to
						'Detailed Inventory for card'
						{
							$InventoryOf = $TempNode1
						}

						'ASM OS License Type'
						{
							if ($InventoryOf -eq 'ASM Module')
							{
								switch ($TempNode1)
								{
									'Srv 2008' { $WindowsFactoryLicence = 'Windows Server 2008 R2' }
									'Srv 2012' { $WindowsFactoryLicence = 'Windows Server 2012 R2' }
								}
							}
						}

						'HW SKU Type'
						{
							switch ($TempNode1)
							{
								'Regular' { $NodeLicenseSKU = 'SBC' }
							}
						}
						'SW Bundled SBC Type'
						{
							switch ($TempNode1)
							{
								'Normal Bundled SBC' { $SoftwareBundled = 'Yes' }
							}
						}
						'ASM_ComputerModel'				{ $ASM_version = $TempNode1 }
						'ASM_WindowsEthernetMainName' 	{ $ASM_WindowsEthernetMainName = $TempNode1 }
						'ASM_WindowsEthernetMainMac' 	{ $ASM_WindowsEthernetMainMac = $TempNode1 }
						# Does the ASM have 2 NICs?
						'ASM_WindowsEthernetSecName' 	{ $ASM_WindowsEthernetSecName = $TempNode1 }
						'ASM_WindowsEthernetSecMac' 	{ $ASM_WindowsEthernetSecMac = $TempNode1 }
					}

					if ($nodeLine -match 'License Information')
					{
						$LicencedForVideo = $false	 		#Initialised here, we'll test for $null in SIP Sig Gp & report accordingly for a T/F/Null value
						$LicencedForForking = $false 		#Initialised here, we'll test for $null in Call Routing table & report accordingly for a T/F/Null value
						$LicencedForTranscoding = $false 	# "
						$TrialLicense = $false
						$BroadSoftFlag = $false
						while ($true)
						{
							$null = $foreach.MoveNext() #Throw away the 'License Information' header
							$TempLine = $foreach.Current
							If ([string]::IsNullOrEmpty($TempLine)) { break }
							# Reformat those with a count:
							$TempLine = [regex]::replace($TempLine, ', Licensed: 0, Available: 0', ':Disabled') # 'SIP Calls' needs special handling if it's disabled
							$TempLine = [regex]::replace($TempLine, ', Licensed: ', ':Yes:')
							$TempLine = [regex]::replace($TempLine, ', Available: ', ':')
							$TempLine = [regex]::replace($TempLine, 'Unlicensed', 'No:0:0')
							If ($TempLine -match '====' ) 		{ continue } # It's a junk line
							If ($TempLine -match 'Trial License'){ $TrialLicense = $true; continue }
							If ($TempLine -match 'License Type'){ continue } # It's not reported on-screen.
							If ($TempLine -match 'CCPE' ) 		{ continue } # I don't know what it is, but it's not reported on-screen (yet)?
							If ($TempLine -match 'OEM code' ) 	{ continue } # I don't know what it is, but it's not reported on-screen (yet)?
							If ($TempLine -match 'Expiration Date' )
							{
								$LicenceExpiration = [regex]::replace($TempLine, 'Expiration Date:', '')
								$LicenceExpiration = $LicenceExpiration.Trim()
								if ($LicenceExpiration -eq 'NA' ) { $LicenceExpiration = 'Not Applicable' }
								continue
							}
							$LicenceLine = $Templine.Split(':')
							if ($LicenceLine[1] -match 'Enabled')	{ $LicenceLine[1] = 'Yes'}
							if ($LicenceLine[1] -match 'Disabled')	{ $LicenceLine[1] = 'No' }
							If ($LicenceLine.Count -eq 2 )
							{
								if ($LicenceLine[1] -eq 'Yes') { $LicenceLine += 'Unlimited'   ; $LicenceLine += 'Unlimited'   ; $LicenceLine += 'None' }
								if ($LicenceLine[1] -eq 'No' ) { $LicenceLine += 'Not Licensed'; $LicenceLine += 'Not Licensed'; $LicenceLine += 'Not Applicable' }
							}
							for ($i = 0; $i -lt $LicenceLine.Count; $i++) { $LicenceLine[$i] = $LicenceLine[$i].Trim() } # Get rid of any leading or trailing spaces
							while ($LicenceLine.Count -le 5)
							{
								$LicenceLine += ''	# Pad out the array for the extra SweLite "Feature Expiration" column. We'll add content as required in the following code
							}
							If (($LicenceLine[0] -match 'WinSer20') -and ($LicenceLine[1] -eq 'No')) { continue } # These ones only show if they're enabled
							if (($LicenceLine[0] -match 'VideoPassthru') -and ($LicenceLine[1] -match 'Yes'))
							{
								$LicencedForVideo = $true
							}
							if (($LicenceLine[0] -match 'Transcoding') -and ($LicenceLine[1] -match 'Yes'))
							{
								$LicencedForTranscoding = $true
							}
							if (($LicenceLine[0] -match 'Forking') -and ($LicenceLine[1] -match 'Yes'))
							{
								$LicencedForForking = $true
							}
							if ($LicenceLine[1] -eq 'Yes') { $LicenceLine[4] = 'None' } # This *might* require changing once I test some more temporary licences.
							#We need to re-title some from their names in the file to their on-screen name:
							switch ($LicenceLine[0])
							{
								'TDM channels'
								{ 	#Old software ~v3, before the separate Port Licence table was added
									$LicenceLine[0] = 'TDM Channels' # Correct to title case
									[int]$LicencedDS1Ports = $LicenceLine[2]
									[int]$LicencedBRIPorts = $LicenceLine[2]
									[int]$LicencedFXSPorts = $LicenceLine[2]
									[int]$LicencedFXOPorts = $LicenceLine[2]
								}
								'SIP channels'
								{
									if ($SweLite)
									{
										$LicenceLine[0] = 'SIP Signaling Sessions'
										if ($TrialLicense)
										{
											#Overwrite any prior values with those applicable to a trial license:
											$LicenceLine[1] = 'Trial License'
											$LicenceLine[2] = '5'
											$LicenceLine[3] = '<Not Available>'
											$LicenceLine[4] = 'None'
										}
									}
									else
									{
										$LicenceLine[0] = 'SIP Calls'
									}
								}
								'SIP registrations'
								{
									$LicenceLine[0] = 'SIP Registrations' # Correct to title case
									if ($SweLite)
									{
										if ($TrialLicense)
										{
											#Overwrite any prior values with those applicable to a trial license:
											$LicenceLine[1] = 'Trial License'
											$LicenceLine[2] = '5'
											$LicenceLine[3] = '<Not Available>'
											$LicenceLine[4] = 'None'
										}
									}
								}
								'DSP channels'
								{
									if ($SweLite)
									{
										$LicenceLine[0] = 'Enhanced Media Sessions with Transcoding'
										if ($TrialLicense)
										{
											#Overwrite any prior values with those applicable to a trial license:
											$LicenceLine[1] = 'Trial License'
											$LicenceLine[2] = '3'
											$LicenceLine[3] = '<Not Available>'
											$LicenceLine[4] = 'None'
										}
									}
									else
									{
										$LicenceLine[0] = 'DSP Resources'
									}
								}
								'SILK Channels'
								{
									$LicenceLine[0] = 'SILK'
									if ($SweLite)
									{
										if ($TrialLicense)
										{
											#Overwrite any prior values with those applicable to a trial license:
											$LicenceLine[1] = 'Trial License'
											$LicenceLine[2] = '3'
											$LicenceLine[3] = '<Not Available>'
											$LicenceLine[4] = 'None'
										}
									}
								}
								'SREC' 			{ $LicenceLine[0] = 'SIP Recording' }
								'RIPR' 			{ $LicenceLine[0] = 'RIP' }
								'Rest' 			{ $LicenceLine[0] = 'REST' }
								'QoEReporting' 	{ $LicenceLine[0] = 'QoE' }
								'AMR-WB codec' 	{ $LicenceLine[0] = 'AMR-WB' }
								'VideoPassthru' { $LicenceLine[0] = 'Video Passthrough' }
								'CVQReporting'  { $LicenceLine[0] = 'SIP VQ Reporting' }
								'Video Call' 	{ $LicenceLine[0] = 'Video Sessions' }
								'High Capacity' { $LicenceLine[0] = 'High Session Capacity Enabled' }
								'Proxy Local SRTP' { $LicenceLine[0] = 'Enhanced Media Sessions without Transcoding' }
								'WinSer2008R2'
								{
									$LicenceLine[0] = 'Additional WS2008R2 ASM License'
								}
								'WinSer2012R2'
								{
									$LicenceLine[0] = 'Additional WS2012R2 ASM License'
								}
								'BroadSoft Subscriber Data'
								{
									$BroadSoftFlag = $true #Unlicensed, the SweLite doesn't have a BroadSoft line in NodeInfo.txt, but still reports 'Not Licensed' on-screen.
								}
								'DS1 ports'
								{
									$LicenceLine[0] = 'DS1 Ports' # Correct to title case
									if ($SWeLite)
									{
										#Special handling for SweLite Beta - may not be needed by GA.
										#Despite being a virtual appliance, the licence reports 'DS1 ports, Licensed: Unlimited, Available: Unlimited'
										#(NB: towards the end I re-test for the $platform and skip the Port Licenses table for the SweLite)
									}
									else
									{
										# Fixup added in 7.0.0C for CCE reporting DS1 port licences in a different format
										if ($LicenceLine[2] -eq "Not Licensed") { $LicenceLine[2] = 0 }
										[int]$LicencedDS1Ports =  $LicenceLine[2]
									}
								}
								'BRI channels' 	{ $LicenceLine[0] = 'BRI Channels'; [int]$LicencedBRIPorts =  $LicenceLine[2]}
								'BRI ports' 	{ $LicenceLine[0] = 'BRI Ports'; 	[int]$LicencedBRIPorts =  $LicenceLine[2]}
								'FXS ports' 	{ $LicenceLine[0] = 'FXS Ports';	[int]$LicencedFXSPorts =  $LicenceLine[2]}
								'FXO ports' 	{ $LicenceLine[0] = 'FXO Ports';	[int]$LicencedFXOPorts =  $LicenceLine[2]}
							}
							If (($LicenceLine[1] -match 'No') -and ($LicenceLine[2] -match '0')) { $LicenceLine[2] = 'Not Licensed'; $LicenceLine[3] = 'Not Licensed'; $LicenceLine[4] = 'Not Applicable' }
							#Now discard all those that the SWeLite doesn't report:
							if ($SweLite)
							{
								if (('SBA', 'Active Directory', 'Transcoding', 'REST', 'CAS', 'CDR', 'OSPF', 'RIP', 'IPsec', 'RBA', 'QoE', 'Video Passthrough', 'Additional WS2008R2 ASM License', 'Additional WS2012R2 ASM License', 'SIP VQ Reporting') -match $LicenceLine[0]) { continue }
								if ($releaseVersion -ge 9)
								{
									#Forking is now free - no longer licenced, and no longer in the licence table (SweLite Only)
									if (('Forking') -match $LicenceLine[0]) { continue }
								}
							}
							
							#Now bung the values into one of two separate licence tables:
							switch ($LicenceLine[0])
							{
								{($_ -eq 'DS1 Ports') -or ($_ -eq 'BRI Channels') -or ($_ -eq 'BRI Ports') -or ($_ -eq 'FXS Ports') -or ($_ -eq 'FXO Ports') }
								{
									if ($LicenceLine[2] -ne 'Not Licensed') # Don't write port types that are not licenced
									{
										$LicenceObject = @($LicenceLine[0], $LicenceLine[1], $LicenceLine[2], '')
										$PortLicenceCollection += , $LicenceObject
									}
								}
								default
								{
									if ($SweLite)
									{
										$LicenceObject = @($LicenceLine[0], $LicenceLine[1], $LicenceLine[2], $LicenceLine[3], $LicenceLine[4])
									}
									else
									{
										$LicenceObject = @($LicenceLine[0], $LicenceLine[1], $LicenceLine[2], $LicenceLine[3])
									}
									$FeatureLicenceCollection += , $LicenceObject
								}
							}
						} # end "While forever"
						#Did BroadSoft appear in the table?
						if (!$BroadSoftFlag)
						{
							$LicenceObject = @('BroadSoft Subscriber Data', 'No', 'Not Licensed', 'Not Licensed', 'Not Applicable')
							$FeatureLicenceCollection += , $LicenceObject
						}
					}
				}
				#If there are no hardware ports licenced, stamp '-- Table is empty --'
				if ($PortLicenceCollection.Count -eq 0)
				{
					$LicenceObject = @('', '', '-- Table is empty --', '')
					$PortLicenceCollection += , $LicenceObject
				}
				#Reorder the features:
				if ($PSVersion -gt '2')
				{
					#PS v2 won't let me use IndexOf on an Object, and I've not been able to get the following to work in any other format.
					# So if you're running windows 7, you get the licences as they come out of the config file, rather than how they show on screen.
					if ($SweLite)
					{
						$customList = 'SIP Signaling Sessions', 'High Session Capacity Enabled', 'Enhanced Media Sessions with Transcoding', 'Enhanced Media Sessions without Transcoding', 'Video Sessions',  'SIP Registrations', 'Forking',
							'BroadSoft Subscriber Data', 'AMR-WB','SILK'
					}
					else
					{
						$customList = 'TDM Channels', 'SIP Calls', 'SIP Signaling Sessions', 'SIP Registrations', 'DSP Resources', 'SIP Media Sessions', 'Forking', 'SBA', 'Active Directory', 'Transcoding', 'REST', 'CAS',
							'CDR', 'OSPF', 'RIP', 'IPsec', 'RBA', 'QoE', 'BroadSoft Subscriber Data', 'AMR-WB', 'Video Passthrough', 'Additional WS2008R2 ASM License', 'Additional WS2012R2 ASM License', 'SIP VQ Reporting',
							'SILK'
					}
					$FeatureLicenceCollection = $FeatureLicenceCollection | Sort-Object -Property {
						$rank=$customList.IndexOf($_[0]);
						if($rank -ne -1){$rank}
						else{[Double]::PositiveInfinity}
					},[0]
				}
				# Now paste in the last bit:
				if ($LicenceExpiration -eq '') { $LicenceExpiration = '<Not Available>' }
				if (!$SweLite)
				{
					$FeatureLicenceCollection += , @('','', '','') #Blank line separator
					$LicenceObject = @(('Software Bundled: {0}' -f ($SoftwareBundled)), ('Node Licence SKU: {0}' -f ($NodeLicenseSKU)), ('License Expiration: {0}' -f ($LicenceExpiration)), ('Windows Factory License: {0}' -f ($WindowsFactoryLicence)),'')
					$FeatureLicenceCollection += , $LicenceObject
				}

				#If by this point these values haven't changed, they weren't referenced in the file:
				if ($ASM_version -eq 'Not Available') { $ASM_version = 'ASM Module not present' }
				if ($ASM_WindowsEthernetMainName -eq 'Not Available') { $ASM_WindowsEthernetMainName = 'None' }
				if ($ASM_WindowsEthernetMainMac  -eq 'Not Available') { $ASM_WindowsEthernetMainMac  = 'None' }
				if ($ASM_WindowsEthernetSecName  -eq 'Not Available') { $ASM_WindowsEthernetSecName  = 'None' }
				if ($ASM_WindowsEthernetSecMac   -eq 'Not Available') { $ASM_WindowsEthernetSecMac   = 'None' }
			}

			$config  = $xml.selectnodes('/Configuration/Token')
			#---------------------------------------------------------------------------------------------------------
			# We read through the file in sequence building arrays to write to the document as we go. There are a few
			# tables that reference 'future values' - ones that we haven't read from the XML yet.
			# To get around this, we do a quick initial run through the document NOW to extract those objects we need
			# to know about out of sequence.
			#---------------------------------------------------------------------------------------------------------
			write-verbose -message 'Initial parse of the config file'
			write-progress -id 1 -Activity 'Initialising' -Status 'Initial parse of the config file' -PercentComplete (6)
			ForEach($node in $config)
			{
				$id = $node.getAttribute('name')

				switch ($id)
				{
					# ---- Node Hardware - Logical Interfaces ----------
					# This code addresses Sonus bug SYM-16535/SYM-18340, where the same NIC exists in the config file twice - an illegal config.
					# It reads all of the logical interfaces into an array, then checks for dupes. It discards the older of the dupes, marking
					# this i/f ID as bad for the later code to discard.
					'NodeHardware'
					{
						$NodeHardwareGroups = $node.GetElementsByTagName('Token')
						ForEach($NodeHardwareGroup in $NodeHardwareGroups)
						{
							if ($NodeHardwareGroup.name -eq 'LogicalInterfaces')
							{
								$LogicalInterfacesList = @()
								$BadLogicalInterfaces = @()
								$LogicalInterfaces = $NodeHardwareGroup.GetElementsByTagName('ID')
								if ($LogicalInterfaces.Count -ne 0)
								{
									ForEach ($LogicalInterface in $LogicalInterfaces)
									{
										if ($LogicalInterface.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($LogicalInterface.IE.classname -eq 'NETSERV_INTERFACE_CFG_IE')
										{
											foreach ($SingleInterface in $LogicalInterfacesList)
											{
												if ($SingleInterface[2] -match $LogicalInterface.IE.ifName)
												{
													write-verbose -message ('Duplicate logical interface {0} flagged for discard' -f $LogicalInterface.IE.ifName)
													if ($LogicalInterface.IE.SystemRelease -gt $SingleInterface[1])
													{
														$BadLogicalInterfaces += $SingleInterface[0]
													}
													else
													{
														$BadLogicalInterfaces += $LogicalInterface.value
													}
												}
											}
											$LogicalInterfacesList += ,($LogicalInterface.value, $LogicalInterface.IE.SystemRelease, $LogicalInterface.IE.ifName)
										}
									}
								}
							}
							#---- Network Adapter ----------
							if ($NodeHardwareGroup.name -eq 'NetworkAdapter')
							{
								#---- Bridge & VLAN ----------
								$BridgeValues = $NodeHardwareGroup.GetElementsByTagName('Token')
								if ($BridgeValues.Count -ne 0)
								{
									ForEach ($BridgeValue in $BridgeValues)
									{
										if ($BridgeValue.name -eq 'Bridge')
										{
											$BridgeGroups = $BridgeValue.GetElementsByTagName('ID')
											ForEach ($BridgeGroup in $BridgeGroups)
											{
												if ($BridgeGroup.IE.classname -eq $null) { continue } # Empty / deleted entry
												if ($BridgeGroup.IE.classname -eq 'NETSERV_VLAN_CFG_IE')
												{
													if ($BridgeGroup.IE.vlan_id -ge 4000)
													{
														$NetworkVLANLookup.Add($BridgeGroup.IE.vlan_id,((Fix-NullDescription -TableDescription $BridgeGroup.IE.vlan_description -TableValue $BridgeGroup.IE.vlan_id -TablePrefix 'VLAN ')))
													}
													else
													{
														$NetworkVLANLookup.Add($BridgeGroup.IE.vlan_id,('VLAN ' + $BridgeGroup.IE.vlan_id + ' - ' + (Fix-NullDescription -TableDescription $BridgeGroup.IE.vlan_description -TableValue $BridgeGroup.IE.vlan_id -TablePrefix 'VLAN ')))
													}
												}

												if ($BridgeGroup.IE.classname -eq 'NETSERV_MSTP_INSTANCE_CFG_IE')
												{
													$MSTPInstanceNameLookup.Add(($BridgeGroup.value), (Fix-NullDescription -TableDescription $BridgeGroup.IE.Description -TableValue $BridgeGroup.IE.mstpInstance -TablePrefix 'MST Instance ')) # Referenced by Node Interfaces / Bridge / VLAN
													$MSTPInstanceIdLookup.Add(($BridgeGroup.value), $BridgeGroup.IE.mstpInstance ) # Referenced by Node Interfaces / Bridge / VLAN
												}
											}
										}
									}
								}
							}
						}
					}
					#---------- IP Routing ------------------------
					'IPRouting'
					{
						$IPRoutingGroups = $node.GetElementsByTagName('Token')
						ForEach($IPRoutingGroup in $IPRoutingGroups)
						{
							# ---- IPv4 ACL's ----------
							if ($IPRoutingGroup.name -eq 'AccessCntrlList')
							{
								$ACLGroups = $IPRoutingGroup.GetElementsByTagName('ID')
								ForEach ($ACLGroup in $ACLGroups)
								{
									if ($ACLGroup.IE.classname -eq $null) { continue } # Empty / deleted entry
									if (($ACLGroup.IE.classname -eq 'CONFIG_DESCRIPTION_IE') -or ($ACLGroup.IE.classname -eq 'ACCESS_CONTROL_LIST_IE')) #The latter is new in Rls 3.0
									{
										$ACLTableLookup.Add($ACLGroup.value, (Fix-NullDescription -TableDescription $ACLGroup.IE.Description -TableValue $ACLGroup.value -TablePrefix 'ACL #')) # Referenced by Ports / Logical Interfaces
									}
								}
							}
							# ---- IPv6 ACL's ----------
							if ($IPRoutingGroup.name -eq 'IPv6AccessCntrlList')
							{
								$ACLGroups = $IPRoutingGroup.GetElementsByTagName('ID')
								ForEach ($ACLGroup in $ACLGroups)
								{
									if ($ACLGroup.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($ACLGroup.IE.classname -eq 'ACCESS_CONTROL_LIST_IE')
									{
										$ACLv6TableLookup.Add($ACLGroup.value, (Fix-NullDescription -TableDescription $ACLGroup.IE.Description -TableValue $ACLGroup.value -TablePrefix 'ACL #')) # Referenced by Ports / Logical Interfaces
									}
								}
							}
						}
					}

					#---------- System ------------------------
					'SBA'
					{
						$SBAGroups = $node.GetElementsByTagName('Token')
						ForEach ($SBAGroup in $SBAGroups)
						{
							# ---- Skype / Lync CAC Profiles ----------
							if ($SBAGroup.name -eq 'CACProfile')
							{
								$CACProfiles = $node.GetElementsByTagName('ID')
								ForEach ($CACProfile in $CACProfiles)
								{
									if ($CACProfile.IE.classname -eq 'CAC_PROFILE_CONFIG')
									{
										$CACProfileLookup.Add($CACProfile.value, (Fix-NullDescription -TableDescription $CACProfile.IE.Description -TableValue $CACProfile.value -TablePrefix 'CAC Profile #')) # Referenced by Network Monitoring / Link Monitors
									}
								}
							}
						}
					}

					#---------- System ------------------------
					'System'
					{
						$systemgroups = $node.GetElementsByTagName('Token')
						ForEach ($systemgroup in $systemgroups)
						{
							# ---- SystemNet ----------
							if ($systemgroup.name -eq 'SystemNet')
							{
								$SystemNodeLevelData = @()
								if ($systemgroup.IE.classname -eq 'SYSTEM_SERV_NET_CFG_IE')
								{
									if ($NodeInfo -ne '')
									{
										if ($NodeHostname -eq '')
										{
											write-warning -Message ("Hostname ""$($NodeHostname)"" from NodeInfo <> hostname ""$($systemgroup.IE.NodeName)"" from symphonyconfig.")
										}
										elseif ($NodeHostname.CompareTo($systemgroup.IE.NodeName) -ne 0)
										{
											write-warning -Message ("Hostname ""$($NodeHostname)"" from NodeInfo <> hostname ""$($systemgroup.IE.NodeName)"" from symphonyconfig. Aborting")
											throw "The Node's hostname in NodeInfo differs to the hostname in the symphonyconfig file"
										}
									}
								}
							}
							# ---- Certificates ----------
							if ($systemgroup.name -eq 'Certificates')
							{
								$SbcCertificates = $systemgroup.GetElementsByTagName('ID')
								if ($SbcCertificates.Count -ne 0)
								{
									ForEach ($SbcCertificate in $SbcCertificates)
									{
										if ($SbcCertificate.IE.classname -eq $null) { continue } # Empty / deleted entry
										if (($SbcCertificate.IE.classname -eq 'CERTIFICATE_FILE_DATA_CFG_IE') -and (($SbcCertificate.IE.CertFileType -eq '1') -or ($SbcCertificate.IE.CertFileType -eq '4')))
										{
											#CertFileType: 1 = user-provided SBC cert, 4 = Ribbon default cert
											$CertCommonName = "Unknown"
											switch ($SbcCertificate.IE.CertCommonName)
											{
												{($_ -eq 'sbc1000') -or ($_ -eq 'sbc2000')}
												{
													$CertCommonName = "SBC Edge Certificate"
												}
												default
												{
													$CertCommonName = $SbcCertificate.IE.CertCommonName
												}
											}
											$CertificateLookup.Add($SbcCertificate.value, $CertCommonName) # Referenced by Security - TLS Profiles
										}
									}
								}
							}
							# ---- SupplementCertificates ----------
							if ($systemgroup.name -eq 'SupplementCertificates')
							{
								$SbcCertificates = $systemgroup.GetElementsByTagName('ID')
								if ($SbcCertificates.Count -ne 0)
								{
									ForEach ($SbcCertificate in $SbcCertificates)
									{
										if ($SbcCertificate.IE.classname -eq $null) { continue } # Empty / deleted entry
										if (($SbcCertificate.IE.classname -eq 'CERTIFICATE_FILE_DATA_CFG_IE') -and ($SbcCertificate.IE.CertFileType -eq '5'))
										{
											$SbcCertificateValue = $SbcCertificate.Value
											if ($SbcCertificateValue -ge 1000) { $SbcCertificateValue -= 1000 } #Remove the 1000 offset
											$CertificateLookup.Add([string]$SbcCertificateValue, $SbcCertificate.IE.CertCommonName) # Referenced by Security - TLS Profiles
										}
									}
								}
							}
						}
					}

					#---------- REMOTE LOG SERVERS ------------
					'Logging'
					{
						$LoggingElements = $node.GetElementsByTagName('Token')
						ForEach($LoggingElement in $LoggingElements)
						{
							# ---- Syslog Servers ----------
							if ($LoggingElement.name -eq 'SyslogServer')
							{
								$RemoteLogServers = $LoggingElement.GetElementsByTagName('ID')
								ForEach ($RemoteLogServer in $RemoteLogServers)
								{
									if ($RemoteLogServer.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($RemoteLogServer.IE.classname -eq 'LOGGER_SYSLOG_DEST')
									{
										$LogServerLookup.Add($RemoteLogServer.value, $RemoteLogServer.IE.ServerAddress) # Referenced by Logging Configuration - Subsystems
									}
								}
							}
						}
					}

					#---------- IP SEC ------------
					'IPSec'
					{
						$IPSecGroups = $node.GetElementsByTagName('Token')
						ForEach($IPSecGroup in $IPSecGroups)
						{
							# ---- Tunnel Tables ----------
							if ($IPSecGroup.name -eq 'IPSecConnectionTable')
							{
								$IPSecTunnels = $IPSecGroup.GetElementsByTagName('ID')
								ForEach ($IPSecTunnel in $IPSecTunnels)
								{
									# ---- Syslog Servers ----------
									if ($IPSecTunnel.IE.classname -eq 'NETSERV_IPSEC_CONNECTION_OBJ_CFG_IE')
									{
										$IPSecTunnelLookup.Add($IPSecTunnel.value, $IPSecTunnel.IE.TunnelName) # Referenced by Network Monitoring / Link Monitors
									}
								}
							}
						}
					}

					#---------- SIGNALING GROUPS ------------
					'SignalingGroups'
					{
						$SignalingGroups = $node.GetElementsByTagName('ID')
						ForEach ($SignalingGroup in $SignalingGroups)
						{
							if ($SignalingGroup.IE.classname -eq $null) { continue } # Empty / deleted entry
							if (($SignalingGroup.IE.classname -eq 'ISDN_SG_PROFILE_CONFIG_ENTRY_IE') -or ($SignalingGroup.IE.classname -eq 'SIP_CFG_SG_COMMON_IE') -or ($SignalingGroup.IE.classname -eq 'CAS_SG_CFG_IE'))
							{
								if ([int]$SignalingGroup.value -lt 50000)
								{
									$SgTableLookup.Add($SignalingGroup.value, ('{0}' -f (Fix-NullDescription -TableDescription $SignalingGroup.IE.Description -TableValue $SignalingGroup.value -TablePrefix 'Signaling Group Table #')))
								}
								else
								{
									$SgTableLookup.Add($SignalingGroup.value, ('{0}' -f (Fix-NullDescription -TableDescription $SignalingGroup.IE.Description -TableValue $SignalingGroup.value -TablePrefix 'SIPREC SG #'))) # Referenced by Signaling Groups (from v9.0.0+)
								}
							}
						}
					}

					#---------- ACTION SETS ------------
					'ActionSets'
					{
						$ActionSets = $node.GetElementsByTagName('ID')
						ForEach ($ActionSet in $ActionSets)
						{
							if ($ActionSet.IE.classname -eq $null) { continue } # Empty / deleted entry
							if ($ActionSet.IE.classname -eq 'CR_ACTION_SET_TABLE_LIST_CONFIG')
							{
								$ActionSetLookup.Add($ActionSet.value, (Fix-NullDescription -TableDescription $ActionSet.IE.Description -TableValue $ActionSet.value -TablePrefix 'Action Set Table #')) # Referenced by Signaling Groups
							}
						}
					}

					#---------- TONE TABLES ------------------------
					'ToneTables'
					{
						$ToneTables = $node.GetElementsByTagName('ID')
						ForEach ($ToneTable in $ToneTables)
						{
							if ($ToneTable.IE.classname -eq $null) { continue } # Empty / deleted entry
							if ($ToneTable.IE.classname -eq 'CONFIG_DESCRIPTION_IE')
							{
								$ToneTableLookup.Add($ToneTable.value, (Fix-NullDescription -TableDescription $ToneTable.IE.Description -TableValue $ToneTable.value -TablePrefix 'Tone Table #')) # Referenced by Signaling Groups
							}
						}
					}

					#---------- CALL ROUTES ------------
					'CallRouting'
					{
						$callroutes = $node.GetElementsByTagName('ID')
						ForEach ($callroute in $callroutes)
						{
							if ($callroute.IE.classname -eq $null) {continue } # Empty / deleted entry
							if ($callroute.IE.classname -eq 'CR_ROUTE_TABLE_LIST_CONFIG')
							{
								$CallRoutingTableLookup.Add($callroute.value, (Fix-NullDescription -TableDescription $callroute.IE.Description -TableValue $callroute.value -TablePrefix 'Call Route Table #')) # Referenced by ActionConfig & Signaling Groups
							}
						}
					}

					#---------- TIME OF DAY TABLES ------------
					'TimeOfDay'
					{
						$TODTables = $node.GetElementsByTagName('ID')
						ForEach ($TODTable in $TODTables)
						{
							if ($TODTable.IE.classname -eq $null) {continue } # Empty / deleted entry
							if ($TODTable.IE.classname -eq 'TIME_OF_DAY_TABLE')
							{
								$TODTablesLookup.Add($TODTable.value, (Fix-NullDescription -TableDescription $TODTable.IE.Description -TableValue $TODTable.value -TablePrefix 'Call Route Table #')) # Referenced by ActionConfig & Signaling Groups
							}
						}
					}

					#---------- SIP ------------------------
					'SIP'
					{
						$SIPGroups = $node.GetElementsByTagName('Token')
						ForEach($SIPGroup in $SIPGroups)
						{
							# ---- SIP Servers ----------
							if ($SIPGroup.name -eq 'SIPServers')
							{
								$SIPServerGroups = $SIPGroup.GetElementsByTagName('ID')
								ForEach ($SIPServerGroup in $SIPServerGroups)
								{
									if ($SIPServerGroup.IE.classname -eq $null) { continue } # Empty / deleted entry
									if (($SIPServerGroup.IE.classname -eq 'CONFIG_DESCRIPTION_IE') -or ($SIPServerGroup.IE.classname -eq 'SIP_CFG_SERVER_TABLE_LIST_IE')) #The latter is new in Rls 3.0
									{
										$SIPServerTablesLookup.Add($SIPServerGroup.value, (Fix-NullDescription -TableDescription $SIPServerGroup.IE.Description -TableValue $SIPServerGroup.value -TablePrefix 'SIP Server Table #')) # Referenced by SIP Signaling Groups
									}
								}
							}

							# ---- SIP Registrars ----------
							if ($SIPGroup.name -eq 'SIPRegistrars')
							{
								$SIPRegistrars = $SIPGroup.GetElementsByTagName('ID')
								ForEach ($SIPRegistrar in $SIPRegistrars)
								{
									if ($SIPRegistrar.IE.classname -eq $null) { continue } # Empty / deleted entry
									if (($SIPRegistrar.IE.classname -eq 'CONFIG_DESCRIPTION_IE') -or ($SIPRegistrar.IE.classname -eq 'SIP_CFG_REGISTRAR_IE')) #The latter is new in Rls 3.0
									{
										$SIPRegistrarsLookup.Add($SIPRegistrar.value, (Fix-NullDescription -TableDescription $SIPRegistrar.IE.Description -TableValue $SIPRegistrar.value -TablePrefix 'Local Registrar #')) # Referenced by SIP Signaling Groups
									}
								}
							}

							# ---- SIP Registration Tables ----------
							if ($SIPGroup.name -eq 'SIPRegistrationTables')
							{
								$SIPRegistrationTables = $SIPGroup.GetElementsByTagName('ID')
								ForEach ($SIPRegistrationTable in $SIPRegistrationTables)
								{
									if ($SIPRegistrationTable.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($SIPRegistrationTable.IE.classname -eq 'CONFIG_DESCRIPTION_IE')
									{
										$SipRegistrationTableLookup.Add($SIPRegistrationTable.value, (Fix-NullDescription -TableDescription $SIPRegistrationTable.IE.Description -TableValue $SIPRegistrationTable.value -TablePrefix 'Contact Registrant Table #'))
									}
								}
							}

							# ---- SIP AuthoriSation Tables ----------
							if ($SIPGroup.name -eq 'SIPAuthorizationTables')
							{
								$SIPAuthorisationTables = $SIPGroup.GetElementsByTagName('ID')
								ForEach ($SIPAuthorisationTable in $SIPAuthorisationTables)
								{
									if ($SIPAuthorisationTable.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($SIPAuthorisationTable.IE.classname -eq 'CONFIG_DESCRIPTION_IE')
									{
										$SIPAuthorisationTableLookup.Add($SIPAuthorisationTable.value, (Fix-NullDescription -TableDescription $SIPAuthorisationTable.IE.Description -TableValue $SIPAuthorisationTable.value -TablePrefix 'Local Pass-thu Authorization Table #'))
									}
								}
							}

							# ---- SIP Registrars ----------
							if ($SIPGroup.name -eq 'SIPNatPrefixTable')
							{
								$SIPNATPrefixes = $SIPGroup.GetElementsByTagName('ID')
								ForEach ($SIPNATPrefix in $SIPNATPrefixes)
								{
									if ($SIPNATPrefix.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($SIPNATPrefix.IE.classname -eq 'SIP_CFG_NAT_PREFIX_TABLE_LIST_IE')
									{
										$SIPNATPrefixesLookup.Add($SIPNATPrefix.value, (Fix-NullDescription -TableDescription $SIPNATPrefix.IE.Description -TableValue $SIPNATPrefix.value -TablePrefix 'Local Registrar #')) # Referenced by SIP Signaling Groups
									}
								}
							}

							# ---- SIP Credentials Tables ----------
							if ($SIPGroup.name -eq 'SIPCredentialsTables')
							{
								$SIPCredentialsTables = $SIPGroup.GetElementsByTagName('ID')
								ForEach ($SIPCredentialsTable in $SIPCredentialsTables)
								{
									if ($SIPCredentialsTable.IE.classname -eq $null) { continue } # Empty / deleted entry
									if (($SIPCredentialsTable.IE.classname -eq 'SIP_CFG_USER_CREDENTIALS_LIST_IE') -or ($SIPCredentialsTable.IE.classname -eq 'CONFIG_DESCRIPTION_IE')) #Changed to the former in Rls 3.
									{
										$SipCredentialsTableLookup.Add($SIPCredentialsTable.value, (Fix-NullDescription -TableDescription $SIPCredentialsTable.IE.Description -TableValue $SIPCredentialsTable.value -TablePrefix 'Remote Authorization Table #'))
									}
								}
							}

							# ---- SIP Profiles ----------
							if ($SIPGroup.name -eq 'Profile')
							{
								$SIPProfiles = $SIPGroup.GetElementsByTagName('ID')
								ForEach ($SIPProfile in $SIPProfiles)
								{
									if ($SIPProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($SIPProfile.IE.classname -eq 'SIP_CFG_PROFILE_IE')
									{
										$SipProfileIdLookup.Add($SIPProfile.value, (Fix-NullDescription -TableDescription $SIPProfile.IE.Description -TableValue $SIPProfile.value -TablePrefix 'SIP Profile #')) # Referenced by SIP Signaling Groups
									}
								}
							}

							# ---- SIP Message Manipulation Rules ----------
							if ($SIPGroup.name -eq 'SIPSPRMessageRules')
							{
								$SIPMessageRules = $SIPGroup.GetElementsByTagName('ID')
								ForEach ($SIPMessageRule in $SIPMessageRules)
								{
									if ($SIPMessageRule.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($SIPMessageRule.IE.classname -eq 'SPR_MESSAGE_TABLE_IE')
									{
										$SIPMessageRuleLookup.Add($SIPMessageRule.value, (Fix-NullDescription -TableDescription $SIPMessageRule.IE.Description -TableValue $SIPMessageRule.value -TablePrefix 'Message Rule Table #')) # Referenced by SIP Signaling Groups
									}
								}
							}

							# ---- SIP Message Rule Rules - Element Descriptors ----------
							if ($SIPGroup.name -eq 'ElementDescriptor')
							{
								#We do more than pull out the Value and Description. For these guys we pull the entire content, ready to paste into a Message Rule Table.
								$SIPElementDescriptors = $SIPGroup.GetElementsByTagName('ID')
								ForEach ($SIPElementDescriptor in $SIPElementDescriptors)
								{
									if ($SIPElementDescriptor.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($SIPElementDescriptor.IE.classname -eq 'SPR_REPAIR_ELEMENT_DESCRIPTOR_IE')
									{
										$SIPElementContent = $SipElementDescElementClassLookup.Get_Item($SIPElementDescriptor.IE.ElementClass) + ': ' + $SipElementDescActionLookup.Get_Item($SIPElementDescriptor.IE.Action) + "`n"
										if ($SIPElementDescriptor.IE.Name -ne '') { $SIPElementContent += ("Name: {0}`n" -f $SIPElementDescriptor.IE.Name) }
										switch ($SIPElementDescriptor.IE.Type)
										{
											'0' #Literal
											{
												if ($SIPElementDescriptor.IE.Action -ne 3) # "Remove" doesn't show the below
												{
													$SIPElementContent += "Type of Value: Literal`n"
													$SIPElementContent += ("Value: {0}`n" -f $SIPElementDescriptor.IE.Value)
												}
											}
											'1' #Token
											{
												$SIPElementContent += "Type of Value: Token`n"
												$SIPElementContent += ("Value: {0}`n"  -f $SIPElementDescriptor.IE.Value)
												if ($SIPElementDescriptor.IE.Action -ne 4) # "Copy Value To" doesn't offer Prefix/Suffix
												{
													$SIPElementContent += ("Prefix: {0}`n" -f $SIPElementDescriptor.IE.Prefix)
													$SIPElementContent += ("Suffix: {0}`n" -f $SIPElementDescriptor.IE.Suffix)
												}
											}
											'2' #Regex
											{
												$SIPElementContent += "Type of Value: Regex`n"
												$SIPElementContent += ("Match Regex: {0}`n"   -f $SIPElementDescriptor.IE.Value)
												$SIPElementContent += ("Replace Regex: {0}`n" -f $SIPElementDescriptor.IE.RegexReplace)
											}
										}
									}
									$SIPMessageRuleElementLookup.Add($SIPElementDescriptor.value, $SIPElementContent) # Referenced by SIP Msg Rule Tables
								}
							}
						}
					}

					 #---------- Media  ------------
					'Media'
					{
						$systemgroups = $node.GetElementsByTagName('Token')
						ForEach ($systemgroup in $systemgroups)
						{
							# ---- MediaListProfiles ----------
							if ($systemgroup.name -eq 'MediaListProfiles')
							{
								$MediaListProfiles = $systemgroup.GetElementsByTagName('ID')
								ForEach ($MediaListProfile in $MediaListProfiles)
								{
									if ($MediaListProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($MediaListProfile.IE.classname -eq 'MEDIALIST_PROFILE_IE')
									{
										$MediaListProfileLookup.Add($MediaListProfile.value, $MediaListProfile.IE.Description)
									}
								}
							}
							# ---- MediaCryptoProfiles - on-screen as 'SDES-SRTP Profiles' as on v8.0 ----------
							if ($systemgroup.name -eq 'MediaCryptoProfiles')
							{
								$MediaCryptoProfiles = $systemgroup.GetElementsByTagName('ID')
								ForEach ($MediaCryptoProfile in $MediaCryptoProfiles)
								{
									if ($MediaCryptoProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($MediaCryptoProfile.IE.classname -eq 'CRYPTO_PROFILE_IE')
									{
										$SDESMediaCryptoProfileLookup.Add($MediaCryptoProfile.value, $MediaCryptoProfile.IE.Description)
									}
								}
							}
							# ---- DTLS MediaCryptoProfiles - new in v8.0 ----------
							if ($systemgroup.name -eq 'DTLSProfiles')
							{
								$MediaCryptoProfiles = $systemgroup.GetElementsByTagName('ID')
								ForEach ($MediaCryptoProfile in $MediaCryptoProfiles)
								{
									if ($MediaCryptoProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($MediaCryptoProfile.IE.classname -eq 'MSC_CFG_DTLS_PROFILE_IE')
									{
										$DTLSMediaCryptoProfileLookup.Add($MediaCryptoProfile.value, $MediaCryptoProfile.IE.Description)
									}
								}
							}
							# ---- FaxProfiles ----------
							if ($systemgroup.name -eq 'FaxProfiles')
							{
								$faxProfiles = $systemgroup.GetElementsByTagName('ID')
								ForEach ($faxProfile in $faxProfiles)
								{
									if ($faxProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($faxProfile.IE.classname -eq 'FAX_CODEC_PROFILE_IE')
									{
										$VoiceFaxProfilesLookup.Add($faxProfile.value, $faxProfile.IE.Description) # Referenced by Media List Profiles
									}
								}
							}
							# ---- VoiceProfiles ----------
							if ($systemgroup.name -eq 'VoiceProfiles')
							{
								$VoiceProfiles = $systemgroup.GetElementsByTagName('ID')
								ForEach ($VoiceProfile in $VoiceProfiles)
								{
									if ($VoiceProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($VoiceProfile.IE.classname -eq 'VOICE_CODEC_PROFILE_IE')
									{
										$VoiceFaxProfilesLookup.Add($VoiceProfile.value, $VoiceProfile.IE.Description) # Referenced by Media List Profiles
									}
								}
							}
						}
					}
					#---------- Telephony Mapping Tables ------------------------
					'TelephonyMappingTables'
					{
						$TMTGroups = $node.GetElementsByTagName('Token')
						ForEach ($TMTGroup in $TMTGroups)
						{
							# ---- Reroute Table  ----------
							if ($TMTGroup.name -eq 'RerouteTable')
							{
								$RerouteTables = $TMTGroup.GetElementsByTagName('ID')
								ForEach ($RerouteTable in $RerouteTables)
								{
									if ($RerouteTable.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($RerouteTable.IE.classname -eq 'CC_REROUTE_CFG_IE')
									{
										$RerouteTablesLookup.Add($RerouteTable.value, $RerouteTable.IE.Description)
									}
								}
							}
							# ---- MsgTranslationTable ----------
							if ($TMTGroup.name -eq 'MsgTranslationTable')
							{
								$MsgTranslationTables = $TMTGroup.GetElementsByTagName('ID')
								ForEach ($MsgTranslationTable in $MsgTranslationTables)
								{
									if ($MsgTranslationTable.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($MsgTranslationTable.IE.classname -eq 'MESSAGE_PI_TRANSLATION_TABLE_IE')
									{
										$MsgTranslationTablesLookup.Add($MsgTranslationTable.value, $MsgTranslationTable.IE.Description)
									}
								}
							}
							# ---- SIPToQ850Table ----------
							if ($TMTGroup.name -eq 'SIPToQ850')
							{
								$SIPToQ850Tables = $TMTGroup.GetElementsByTagName('ID')
								ForEach ($SIPToQ850Table in $SIPToQ850Tables)
								{
									if ($SIPToQ850Table.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($SIPToQ850Table.IE.classname -eq 'CONFIG_DESCRIPTION_IE')
									{
										$SIPtoQ850TableLookup.Add($SIPToQ850Table.value, $SIPToQ850Table.IE.Description)
									}
								}
							}
							# ---- Q850ToSIPTable ----------
							if ($TMTGroup.name -eq 'Q850ToSIP')
							{
								$Q850ToSIPTables = $TMTGroup.GetElementsByTagName('ID')
								ForEach ($Q850ToSIPTable in $Q850ToSIPTables)
								{
									if ($Q850ToSIPTable.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($Q850ToSIPTable.IE.classname -eq 'CONFIG_DESCRIPTION_IE')
									{
										$Q850ToSIPTableLookup.Add($Q850ToSIPTable.value, $Q850ToSIPTable.IE.Description)
									}
								}
							}
						}
					}
					#---------- Radius ------------------------
					'RAD'
					{
						$RadiusUserGroups = $node.GetElementsByTagName('Token')
						ForEach ($RadiusUserGroup in $RadiusUserGroups)
						{
							# ---- Radius Servers ----------
							if ($RadiusUserGroup.name -eq 'Server')
							{
								$RadiusServers = $RadiusUserGroup.GetElementsByTagName('ID')
								if ($RadiusServers.Count -ne 0)
								{
									ForEach ($RadiusServer in $RadiusServers)
									{
										if ($RadiusServer.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($RadiusServer.IE.classname -eq 'RAD_RADIUS_SERVERS')
										{
											$RadiusServerLookup.Add($RadiusServer.value, $RadiusServer.IE.Description)
										}
									}
								}
							}
						}
					}
					#---------- CAS ------------------------
					'CAS'
					{
						$CASProfileGroups = $node.GetElementsByTagName('Token')
						ForEach ($CASProfileGroup in $CASProfileGroups)
						{
							# ---- CASLoopStartE&M ----------
							if ($CASProfileGroup.name -eq 'CASEnM')
							{
								$CASEnMProfiles = $CASProfileGroup.GetElementsByTagName('ID')
								ForEach ($CASEnMProfile in $CASEnMProfiles)
								{
									if ($CASEnMProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($CASEnMProfile.IE.classname -eq 'CAS_ENM_CFG_IE')
									{
										$CASSignalingProfileLookup.Add($CASEnMProfile.value, $CASEnMProfile.IE.Description)
									}
								}
							}
							# ---- CASLoopStartR2MFC ----------
							if ($CASProfileGroup.name -eq 'CASR2')
							{
								$CASR2Profiles = $CASProfileGroup.GetElementsByTagName('ID')
								ForEach ($CASR2Profile in $CASR2Profiles)
								{
									if ($CASR2Profile.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($CASR2Profile.IE.classname -eq 'CAS_RTWO_CFG_IE')
									{
										$CASSignalingProfileLookup.Add($CASR2Profile.value, $CASR2Profile.IE.Description)
									}
								}
							}
							# ---- CASLoopStartFxs ----------
							if ($CASProfileGroup.name -eq 'CASLoopStartFxs')
							{
								$CASLoopStartFXSProfiles = $CASProfileGroup.GetElementsByTagName('ID')
								ForEach ($CASLoopStartFXSProfile in $CASLoopStartFXSProfiles)
								{
									if ($CASLoopStartFXSProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($CASLoopStartFXSProfile.IE.classname -eq 'CAS_LOOPSTART_FXS_CFG_IE')
									{
										$CASSignalingProfileLookup.Add($CASLoopStartFXSProfile.value, $CASLoopStartFXSProfile.IE.Description)
									}
								}
							}
							# ---- CASLoopStartFXO ----------
							if ($CASProfileGroup.name -eq 'CASLoopStartFxo')
							{
								$CASLoopStartFXOProfiles = $CASProfileGroup.GetElementsByTagName('ID')
								ForEach ($CASLoopStartFXOProfile in $CASLoopStartFXOProfiles)
								{
									if ($CASLoopStartFXOProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($CASLoopStartFXOProfile.IE.classname -eq 'CAS_LOOPSTART_FXO_CFG_IE')
									{
										$CASSignalingProfileLookup.Add($CASLoopStartFXOProfile.value, $CASLoopStartFXOProfile.IE.Description)
									}
								}
							}
							# ---- CASSupplementary ----------
							if ($CASProfileGroup.name -eq 'CASSupplementary')
							{
								$CASSupplementaryProfiles = $CASProfileGroup.GetElementsByTagName('ID')
								ForEach ($CASSupplementaryProfile in $CASSupplementaryProfiles)
								{
									if ($CASSupplementaryProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
									if ($CASSupplementaryProfile.IE.classname -eq 'CAS_SUPPLEMENTARY_CFG_IE')
									{
										$CASSupplementaryProfileLookup.Add($CASSupplementaryProfile.value, $CASSupplementaryProfile.IE.Description)
									}
								}
							}
						}
					}
				}
			}

			#---------------------------------------------------------------------------------------------------------
			#This is the 'real' run through the XML file, where we extract the remainder of what we need & everything is written to Word
			#---------------------------------------------------------------------------------------------------------
			write-verbose -message 'Main	parse of the config file'
			write-progress -id 1 -Activity 'Initialising' -Status 'Main parse of the config file' -PercentComplete (8)

			ForEach($node in $config)
			{
				$id = $node.getAttribute('name')

				switch ($id)
				{
					#---------- Users ------------------------
					'Users'
					{
						$UserGroups = $node.GetElementsByTagName('Token')
						$ADUserGroupCollection = @()
						ForEach ($UserGroup in $UserGroups)
						{
							# ---- AD User Groups ----------
							if ($UserGroup.name -eq 'Groups')
							{
								$ADUserGroups = $UserGroup.GetElementsByTagName('ID')
								$ADUserGroupColumnTitles = @('Group Name', 'Access Level', 'Primary Key')
								if ($ADUserGroups.Count -ne 0)
								{
									ForEach ($ADUserGroup in $ADUserGroups)
									{
										if ($ADUserGroup.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($ADUserGroup.IE.classname -eq 'UAS_USER_GROUP')
										{
											$ADUserGroupObject = @($ADUserGroup.IE.GroupName, $AccountAccessLevelLookup.Get_Item($ADUserGroup.IE.AuthLevel), $ADUserGroup.value)
											$ADUserGroupCollection += , $ADUserGroupObject
										}
									}
									#The $ADUserGroupCollection is added to $SecurityData just prior to printing to ensure the correct ordering in the DOCX
								}
							}
						}
					}

					#---------- Node Hardware ------------------------
					'NodeHardware'
					{
						$NodeHardwareGroups = $node.GetElementsByTagName('Token')
						$LineCardCollection = @()
						ForEach ($NodeHardwareGroup in $NodeHardwareGroups)
						{
							# ---- Line Cards ----------
							if ($NodeHardwareGroup.name -eq 'LineCards')
							{
								if ($LicencedDS1Ports -eq $null) { [int]$LicencedDS1Ports = 99} # If we've not read from the NodeInfo file, assume no restriction - show it all.
								$LineCards = $NodeHardwareGroup.GetElementsByTagName('ID')
								$LineCardColumnTitles = @('Card', 'Port', 'Enabled', 'Port ID', 'Alias', 'Description', 'Port Type', 'Signaling Type', 'DS1 Framing', 'Line Coding', 'CRC')
								if ($LineCards.Count -ne 0)
								{
									ForEach ($LineCard in $LineCards)
									{
										if ($LineCard.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($LineCard.IE.classname -eq 'LINE_CARD_CONFIG_IE')
										{
											$LineCardPorts = $LineCard.GetElementsByTagName('ID')
											ForEach ($LineCardPort in $LineCardPorts)
											{
												if  ($LineCardPort.IE.classname -eq 'PORT_PROFILE_CONFIG_ENTRY_IE')
												{
													$ISDNLineCardPortName = [regex]::replace($LineCardPort.IE.PortName, '-' , ' ') # Turn the dash into a space
													if ($LineCardPort.IE.PortPhysicalType -eq '0')
													{
														#E1
														$LineCardObject = @($LineCard.value, $LineCardPort.value, $EnabledLookup.Get_Item($LineCardPort.IE.Enabled), $ISDNLineCardPortName, $LineCardPort.IE.PortAlias, $LineCardPort.IE.PortDescription, $IsdnPortPhysicalType.Get_Item($LineCardPort.IE.PortPhysicalType), $IsdnPortSignalingType.Get_Item($LineCardPort.IE.PortSignalingType), '<n/a>', $IsdnPortCodingType.Get_Item($LineCardPort.IE.PortCodingType), $EnabledLookup.Get_Item($LineCardPort.IE.PortCRCType))
													}
													else
													{
														#T1
														$LineCardObject = @($LineCard.value, $LineCardPort.value, $EnabledLookup.Get_Item($LineCardPort.IE.Enabled), $ISDNLineCardPortName, $LineCardPort.IE.PortAlias, $LineCardPort.IE.PortDescription, $IsdnPortPhysicalType.Get_Item($LineCardPort.IE.PortPhysicalType), $IsdnPortSignalingType.Get_Item($LineCardPort.IE.PortSignalingType), $IsdnPortDS1Framing.Get_Item($LineCardPort.IE.PortFramingType), $IsdnPortCodingType.Get_Item($LineCardPort.IE.PortCodingType), '<n/a>')
													}
													if ($LicencedDS1Ports -gt 0)
													{
														$LineCardCollection += , $LineCardObject
														$LicencedDS1Ports --
													}
												}
											}
										}
									}
									#$NodeHardwareData += , ('Ports', 'Telephony', $LineCardColumnTitles, $LineCardCollection) - this is now added just prior to writing to Word
								}
							}
							# ---- Network Adapter ----------
							if ($NodeHardwareGroup.name -eq 'NetworkAdapter')
							{
								$NetworkAdapters = $NodeHardwareGroup.GetElementsByTagName('ID')
								if ($NetworkAdapters.Count -ne 0)
								{
									$NodeHardwarePortData = @()
									ForEach ($NetworkAdapter in $NetworkAdapters)
									{
										if ($NetworkAdapter.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($NetworkAdapter.IE.classname -eq 'NETSERV_INTERFACE_CFG_IE')
										{
											#Test for and discard a non-equipped ASM/SBA:
											if ('ASM,Sba1,Sba2' -match $NetworkAdapter.IE.IfName)
											{
												#Very old firmware called it Sba1 and Sba2, since renamed ASM
												if ('Not' -match $ASM_version)
												{
													#This bit will only be true if there's no ASM, or we didn't extract from .tar or have -IncludeNodeInfo
													if (($NetworkAdapter.IE.ifHwAddress -eq '') -or ($NetworkAdapter.IE.ifHwAddress -eq $null)) { continue }
												}
											}
											if ('mgmt1' -match $NetworkAdapter.IE.IfName) { continue } #Skip the Admin port - there's nothing to report.
											$NetworkAdapterObject = @()
											# Now let's build the table:
											$NetworkAdapterObject += ,('SPAN-L', 'Identification/Status', 'SPAN-R', 'Networking')
											#From here I split the page into LH and RH columns to deal with all the possible different combinations.
											$NetworkAdapterL1 = @()
											$NetworkAdapterL2 = @()
											$NetworkAdapterR1 = @()
											$NetworkAdapterR2 = @()
											#LH:
											$NetworkAdapterL1 += 'Primary Key'
											$NetworkAdapterL2 += $NetworkAdapter.value
											$NetworkAdapterL1 += 'Port ID'
											$NetworkAdapterPortID = $NetworkAdapter.IE.IfName
											$NetworkAdapterPortID = [regex]::replace($NetworkAdapterPortID, 'lan', 'Ethernet ')
											$NetworkAdapterL2 += $NetworkAdapterPortID
											$NetworkAdapterL1 += 'Hardware Type'
											$NetworkAdapterL2 += $PortHardwareTypeLookup.Get_Item($NetworkAdapter.IE.ifHwType)
											$NetworkAdapterL1 +=  'I/F Index'
											$NetworkAdapterL2 += $NetworkAdapter.IE.IfInterfaceIndex
											$NetworkAdapterL1 += 'Port Alias'
											$NetworkAdapterL2 += $NetworkAdapter.IE.IfAlias
											$NetworkAdapterL1 += 'Description'
											$NetworkAdapterL2 += $NetworkAdapter.IE.IfDescription
											if ($platform -eq 'SBC 2000')
											{
												$NetworkAdapterL1 += 'Admin State'
												$NetworkAdapterL2 += $EnabledLookup.Get_Item($NetworkAdapter.IE.Enabled)
											}
											$NetworkAdapterL1 += 'Operational Status'
											$NetworkAdapterL2 += '<Not Available>'
											$NetworkAdapterL1 += 'Up/Down Since'
											$NetworkAdapterL2 += '<Not Available>'
											#RH:
											if ($platform -eq 'SBC 2000')
											{
												$NetworkAdapterR1 += 'ACL In'
												$NetworkAdapterR2 +=  Test-ForNull -LookupTable $ACLTableLookup -value $NetworkAdapter.IE.aclInInstanceID
												$NetworkAdapterR1 += 'ACL Forward'
												$NetworkAdapterR2 +=  Test-ForNull -LookupTable $ACLTableLookup -value $NetworkAdapter.IE.aclForwardInstanceID
											}
											$NetworkAdapterR1 += 'Frame Type'
											if (($NetworkAdapter.IE.ifRedundancy -eq $null) -or ($NetworkAdapter.IE.ifRedundancy -eq '0'))
											{
												$NetworkAdapterR2 +=  Test-ForNull -LookupTable $NetworkAdapterFrameTypeLookup -value $NetworkAdapter.IE.IfFrameType
												$FormattedVLANID = [regex]::replace($NetworkAdapter.IE.ifDefaultVid, '1.' , '')
												$TaggedVLANNames = ''
												if ($NetworkAdapter.IE.ifHybridVlan -eq $null)
												{
													$TaggedVLANNames = '<n/a this rls>'
												}
												else
												{
													if ($NetworkAdapter.IE.ifHybridVlan -ne '')
													{
														$TaggedVLAN_list = ($NetworkAdapter.IE.ifHybridVlan).Split(',') #Value in the file is formatted as '1:4,1:17'
														foreach ($TaggedVLAN in $TaggedVLAN_list)
														{
															$TaggedVLAN = [regex]::replace($TaggedVLAN, '1:' , '')
															$TaggedVLANNames += ($NetworkVLANLookup.Get_Item($TaggedVLAN) + "`n")
														}
													}
												}
												$TaggedVLANNames = Strip-TrailingCR -DelimitedString $TaggedVLANNames
												switch ($NetworkAdapter.IE.IfFrameType)
												{
													'0' #All
													{
														$NetworkAdapterR1 += 'Default Untagged VLAN'
														$NetworkAdapterR2 += $NetworkVLANLookup.Get_Item($FormattedVLANID)
														$NetworkAdapterR1 += 'Tagged VLANs'
														$NetworkAdapterR2 += $TaggedVLANNames
													}
													'1' #Untagged
													{
														$NetworkAdapterR1 += 'Default Untagged VLAN'
														$NetworkAdapterR2 += $NetworkVLANLookup.Get_Item($FormattedVLANID)
													}
													'2' #Tagged
													{
														$NetworkAdapterR1 += 'Tagged VLANs'
														$NetworkAdapterR2 += $TaggedVLANNames
													}
												}
											}
											else
											{
												$NetworkAdapterR2 += 'Untagged'
											}
											#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array
											if ($NetworkAdapterL1.Count -ge $NetworkAdapterR1.Count)
											{
												$arrayCount = $NetworkAdapterL1.Count
											}
											else
											{
												$arrayCount = $NetworkAdapterR1.Count
											}
											for ($i = 0; $i -lt $arrayCount; $i++)
											{
												if (($NetworkAdapterL1[$i] -eq ''   ) -and ($NetworkAdapterL2[$i] -eq ''   ) -and ($NetworkAdapterR1[$i] -eq ''   ) -and ($NetworkAdapterR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
												if (($NetworkAdapterL1[$i] -eq $null) -and ($NetworkAdapterL2[$i] -eq $null) -and ($NetworkAdapterR1[$i] -eq $null) -and ($NetworkAdapterR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
												$NetworkAdapterObject += ,($NetworkAdapterL1[$i], $NetworkAdapterL2[$i], $NetworkAdapterR1[$i], $NetworkAdapterR2[$i])
											}

											if ('ASM,Sba1,Sba2' -match $NetworkAdapter.IE.IfName)
											{
												#Skip the Physical/Data Layer & Ethernet Port Redundancy sections for the ASM
											}
											else
											{
												if ($NetworkAdapter.IE.ifRedundancy -ne $null)
												{
													#New from v7.0 - the RH pane changes to 'Ethernet Port Redundancy'
													$NetworkAdapterObject += ,('SPAN-L', 'Physical/Data Layer', 'SPAN-R', 'Ethernet Port Redundancy')
												}
												else
												{
													$NetworkAdapterObject += ,('SPAN-L', 'Physical/Data Layer', 'SPAN-R', 'Spanning Tree')
												}
												#From here I split the page *again* into LH and RH columns to deal with all the possible different combinations.
												$NetworkAdapterL1 = @()
												$NetworkAdapterL2 = @()
												$NetworkAdapterR1 = @()
												$NetworkAdapterR2 = @()
												#LH:
												$NetworkAdapterL1 += 'Configured Speed'
												$NetworkAdapterL2 += $PortSpeedLookup.Get_Item($NetworkAdapter.IE.ifConfiguredSpeed)
												$NetworkAdapterL1 += 'Negotiated Speed'
												$NetworkAdapterL2 += '<Not Available>'
												$NetworkAdapterL1 += 'Configured Duplexity'
												$NetworkAdapterL2 += $PortDuplexityLookup.Get_Item($NetworkAdapter.IE.ifConfiguredDuplexity)
												$NetworkAdapterL1 += 'Negotiated Duplexity'
												$NetworkAdapterL2 += '<Not Available>'
												#RH:
												if ($NetworkAdapter.IE.ifRedundancy -ne $null)
												{
													#New from v7.0 - the RH pane changes to 'Ethernet Port Redundancy'
													$NetworkAdapterR1 += 'Redundancy'
													switch ($NetworkAdapter.IE.ifRedundancy)
													{
														'0'
														{
															$NetworkAdapterR2 += 'None'
														}
														'1'
														{
															$NetworkAdapterR2 += 'MSTP'
															$NetworkAdapterR1 += 'Protocol BPDU Version Rx'
															$NetworkAdapterR2 += '<Not Available>'
															$NetworkAdapterR1 += 'Protocol BPDU Version Tx'
															$NetworkAdapterR2 += '<Not Available>'
														}
														'2'
														{
															$NetworkAdapterR2 += 'Failover'
															$NetworkAdapterR1 += 'Redundant Port'
															$NetworkAdapterR2 += 'Ethernet ' + $NetworkAdapter.IE.ifRedundantPort
															$NetworkAdapterR1 += 'Role'
															$NetworkAdapterR2 += '<Not Available>'
															$NetworkAdapterR1 += 'State'
															$NetworkAdapterR2 += '<Not Available>'
														}
													}
												}
												else
												{
													$NetworkAdapterR1 += 'MSTP State'
													$NetworkAdapterR2 += (Test-ForNull -LookupTable $EnabledLookup -value $NetworkAdapter.IE.ifMstpStatus)
												}
												#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array
												if ($NetworkAdapterL1.Count -ge $NetworkAdapterR1.Count)
												{
													$arrayCount = $NetworkAdapterL1.Count
												}
												else
												{
													$arrayCount = $NetworkAdapterR1.Count
												}
												for ($i = 0; $i -lt $arrayCount; $i++)
												{
													if (($NetworkAdapterL1[$i] -eq ''   ) -and ($NetworkAdapterL2[$i] -eq ''   ) -and ($NetworkAdapterR1[$i] -eq ''   ) -and ($NetworkAdapterR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
													if (($NetworkAdapterL1[$i] -eq $null) -and ($NetworkAdapterL2[$i] -eq $null) -and ($NetworkAdapterR1[$i] -eq $null) -and ($NetworkAdapterR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
													$NetworkAdapterObject += ,($NetworkAdapterL1[$i], $NetworkAdapterL2[$i], $NetworkAdapterR1[$i], $NetworkAdapterR2[$i])
												}
												if ($platform -eq 'SBC 2000')
												{
													$NetworkAdapterObject += ,('Gigabit Timing Mode', (Test-ForNull -LookupTable $NetworkAdapterGigabitTimingLookup -value $NetworkAdapter.IE.gigabitTimingMode), '','')
												}
											}
											$NodeHardwarePortData += ,('Ports', $NetworkAdapterPortID, '', $NetworkAdapterObject)
										}
									}
								}
								#---- Bridge & VLAN ----------
								$BridgeValues = $NodeHardwareGroup.GetElementsByTagName('Token')
								if ($BridgeValues.Count -ne 0)
								{
									$RegionSettingsColumnTitles = @('Bridge Group ID', 'Protocol', 'Region', 'Revision', 'Learning', 'Aging Time', 'Primary Key')
									$RegionSettingsDataCollection = @()

									$VLANColumnTitles = @('VLAN ID', 'VLAN Description', 'MST Instance ID', 'Primary Key')
									$VLANDataObject = @()
									$VLANDataCollection = @()

									$MSTPColumnTitles = @('Instance ID', 'Description', 'Priority', 'Primary Key')
									$MSTPDataObject = @()
									$MSTPDataCollection = @()
									ForEach ($BridgeValue in $BridgeValues)
									{
										if ($BridgeValue.name -eq 'Bridge')
										{
											$RegionSettings = $BridgeValue.GetElementsByTagName('ID')
											Foreach ($RegionSetting in $RegionSettings)
											{
												if ($RegionSetting.IE.classname -eq $null) { continue } # Empty / deleted entry
												if ($RegionSetting.IE.classname -eq 'NETSERV_BRIDGE_CFG_IE')
												{
													$RegionSettingsDataObject = @()
													$MSTPRegion = '<n/a this rls>'
													$MSTPRevision = '<n/a this rls>'
													$MSTPLearning = '<obsolete>'
													$MSTPAgingTime = '<obsolete>'
													if ($RegionSetting.IE.mstpRegion -ne $null) { $MSTPRegion = $RegionSetting.IE.mstpRegion}
													if ($RegionSetting.IE.mstpRevision -ne $null) { $MSTPRevision = $RegionSetting.IE.mstpRevision}
													if ($RegionSetting.IE.Learning -ne $null) { $MSTPLearning = $RegionSetting.IE.Learning} # Present in v2.0.2 b218
													if ($RegionSetting.IE.AgingTime -ne $null) { $MSTPAgingTime = $RegionSetting.IE.AgingTime} # Present in v2.0.2 b218
													$RegionSettingsDataObject = @($BridgeValue.ID.value, $BridgeRegionSettingsProtocolLookup.Get_Item($RegionSetting.IE.Protocol), $MSTPRegion, $MSTPRevision, $MSTPLearning, $MSTPAgingTime, $BridgeValue.ID.value)
													$RegionSettingsDataCollection += , $RegionSettingsDataObject
												}

												if ($RegionSetting.IE.classname -eq 'NETSERV_VLAN_CFG_IE')
												{
													#if ([int]($RegionSetting.IE.vlan_id) -gt 3999) { continue } #Skip all 'out of range' VLANs - they're the internal ones and don't show in the web i/f
													if ($RegionSetting.IE.mstpInstance -eq $null)
													{
														$VLANDataObject = @($RegionSetting.IE.vlan_id, (Fix-NullDescription -TableDescription $RegionSetting.IE.vlan_description -TableValue $RegionSetting.value -TablePrefix 'VLAN '), '<n/a this rls>', $RegionSetting.value)
													}
													else
													{
														$FormattedMstpInstance = ($RegionSetting.IE.mstpInstance).TrimStart('1')  #SBC 1k (as at v4.0.0 b353 firmware - my Lab) stores the mstpInstance as '1.1'
														$FormattedMstpInstance =		 $FormattedMstpInstance.TrimStart(':.') #SBC 2k (as at v3.1.1 b290 firmware - Lasse's sample) stores mstpInstance as '1:1'
														$InternalVLANList = @('4040','4041','4042','4031','4032','4033','4034')
														if ($InternalVLANList -contains $RegionSetting.IE.vlan_id) { continue } # Skip this one as it's internal. It doesn't show in the VLAN list
														$VLANDataObject = @($RegionSetting.IE.vlan_id, (Fix-NullDescription -TableDescription $RegionSetting.IE.vlan_description -TableValue $RegionSetting.value -TablePrefix 'VLAN '), ($MSTPInstanceIdLookup.Get_Item($FormattedMstpInstance) + ': ' + $MSTPInstanceNameLookup.Get_Item($FormattedMstpInstance)), $RegionSetting.value)
													}
													$VLANDataCollection += , $VLANDataObject
												}

												if ($RegionSetting.IE.classname -eq 'NETSERV_MSTP_INSTANCE_CFG_IE')
												{
													$MSTPDataObject = @($RegionSetting.IE.mstpInstance, (Fix-NullDescription -TableDescription $RegionSetting.IE.Description -TableValue $RegionSetting.IE.mstpInstance -TablePrefix 'MST Instance '), $MstpInstanceBridgePriorityLookup.Get_Item($RegionSetting.IE.mstpInstanceBridgePriority), $RegionSetting.value)
													$MSTPDataCollection += , $MSTPDataObject
												}
											}
										}
									}
								}
							}

							# ---- Logical Interfaces ----------
							if ($NodeHardwareGroup.name -eq 'LogicalInterfaces')
							{
								$LogicalInterfaces = $NodeHardwareGroup.GetElementsByTagName('ID')
								if ($LogicalInterfaces.Count -ne 0)
								{
									$NodeHardwareLogicalInterfaceData = @()
									ForEach ($LogicalInterface in $LogicalInterfaces)
									{
										if ($LogicalInterface.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($LogicalInterface.IE.classname -eq 'NETSERV_INTERFACE_CFG_IE')
										{
											if (($LogicalInterface.IE.NetworkMapped -ne $null) -and ($LogicalInterface.IE.NetworkMapped -eq 0))
											{
												# This is an SWE Lite and this NIC isn't mapped - discard it.
												continue
											}
											#if (([int]$LogicalInterface.value - 14000) -ge 0) { continue } #These are the internal VLANs - they don't show in the Web UI
											$LogicalInterfaceObject = @()
											#Here we build the list of IP addresses to their network ports. (Used in SIP Sig Gps)
											#--------------------------------------------------------------------
											#This code addresses Sonus bug SYM-16535/SYM-18340, where the same NIC exists in the config file twice - an illegal config.
											if ($BadLogicalInterfaces -match $LogicalInterface.value)
											{
												write-warning -message ('NIC {0} is represented twice in the config database. The older version of this NIC has been discarded' -f $LogicalInterface.IE.IfName)
												continue
											}
											#--------------------------------------------------------------------
											# Now let's build the table:
											$LogicalInterfaceObject += ,('SPAN-L', 'Identification/Status', '', '')
											$FormattedIfName = [regex]::replace($LogicalInterface.IE.IfName, 'vlan1.' , 'VLAN ') #Neaten the display of the VLANs
											if ($FormattedIfName.StartsWith('VLAN '))  { $FormattedIfName  = $FormattedIfName + ' IP'}
											switch ($LogicalInterface.IE.IfName)
											{
												'eth0'	   { $FormattedIfName = 'Admin IP' } # SBC 2k
												'local0'	 { $FormattedIfName = 'Loopback 1' }
												'local1'	 { $FormattedIfName = 'Loopback 2' }
												'local2'	 { $FormattedIfName = 'Loopback 3' }
												'local3'	 { $FormattedIfName = 'Loopback 4' }
												'local4'	 { $FormattedIfName = 'Loopback 5' }
												'vlan1.4040' { $FormattedIfName = 'Ethernet 1 IP' } # SBC 1k
												'vlan1.4041' { $FormattedIfName = 'Ethernet 2 IP' } # SBC 1k
												'vlan1.4042' { $FormattedIfName = 'Ethernet 3 IP' } # SBC 1k - 3rd NIC added in hardware refresh
												'vlan1.4031' { $FormattedIfName = 'Ethernet 1 IP' } # SBC 2k
												'vlan1.4032' { $FormattedIfName = 'Ethernet 2 IP' } # SBC 2k
												'vlan1.4033' { $FormattedIfName = 'Ethernet 3 IP' } # SBC 2k
												'vlan1.4034' { $FormattedIfName = 'Ethernet 4 IP' } # SBC 2k
												'mgt0' 		 { $FormattedIfName = 'Admin IP'	  } # SWe Lite
												'Ethernet 1' { $FormattedIfName = 'Ethernet 1 IP' } # SWe Lite
												'Ethernet 2' { $FormattedIfName = 'Ethernet 2 IP' } # SWe Lite
												'Ethernet 3' { $FormattedIfName = 'Ethernet 3 IP' } # SWe Lite
												'Ethernet 4' { $FormattedIfName = 'Ethernet 4 IP' } # SWe Lite
											}
											$LogicalInterfaceObject += ,('Interface Name', $FormattedIfName, '', '')
											$LogicalInterfaceObject += ,('I/F Index', $LogicalInterface.IE.ifInterfaceIndex, '', '')
											$LogicalInterfaceObject += ,('Alias', $LogicalInterface.IE.IfAlias, '', '')
											$LogicalInterfaceObject += ,('Description', $LogicalInterface.IE.IfDescription, '', '')
											$LogicalInterfaceObject += ,('Admin State', $EnabledLookup.Get_Item($LogicalInterface.IE.Enabled), '', '')
											$LogicalInterfaceObject += ,('SPAN-L', 'Networking', '', '')
											$LogicalInterfaceObject += ,('MAC Address', '<Not Available>', '', '')
											if ($LogicalInterface.IE.ipAddressingMode -eq $null)
											{
												$LogicalInterfaceIpAddressingMode = '0' # Force 0 & treat as IPv4 if this SBC is < v6
											}
											else
											{
												$LogicalInterfaceIpAddressingMode = $LogicalInterface.IE.ipAddressingMode
											}

											if ($FormattedIfName -match 'Loopback ')
											{
												# We don't show /offer 'IP Addressing Mode' for the Loopback interfaces
											}
											else
											{
												$LogicalInterfaceObject += ,('IP Addressing Mode', $IpVersionLookup.Get_Item($LogicalInterfaceIpAddressingMode), '', '')
											}
											if ($LogicalInterfaceIpAddressingMode -ne '1')
											{
												#IPv4 or 'both'.
												$LogicalInterfaceObject += ,('SPAN-L', 'IPv4 Information', '', '')
												if ($LogicalInterface.IE.ifIpAddrAssignMethod -eq $null) #New in v3 apparently - Static or DHCP?
												{
													$LogicalAssignmentMethod = $LogicalAddressIPAssignmentMethodLookup.Get_Item($LogicalInterface.IE.ipAddressingMode)
												}
												else
												{
													$LogicalAssignmentMethod = $LogicalAddressIPAssignmentMethodLookup.Get_Item($LogicalInterface.IE.ifIpAddrAssignMethod)
												}
												if ($FormattedIfName -ne 'Admin IP')
												{
													if ($SWeLite)
													{
														#Suppress ACLs - they don't show here at all in the SWe Lite - the NIC r'ship is now set in the ACLs themselves
													}
													else
													{
														$LogicalInterfaceObject += ,('ACL In', (Test-ForNull -LookupTable $ACLTableLookup -value $LogicalInterface.IE.aclInInstanceID), '', '')
														$LogicalInterfaceObject += ,('ACL Out', (Test-ForNull -LookupTable $ACLTableLookup -value $LogicalInterface.IE.aclOutInstanceID), '', '')
														$LogicalInterfaceObject += ,('ACL Forward', (Test-ForNull -LookupTable $ACLTableLookup -value $LogicalInterface.IE.aclForwardInstanceID), '', '')
													}
												}
												if ((($FormattedIfName -match 'Loopback ') -or ($FormattedIfName -eq 'Admin IP')) -and (!$SWeLite))
												{
													# We don't show /offer 'IP Assign Method' for the Admin IP or Loopback interfaces - UNLESS it's a SWe Lite's Admin NIC, in which case we do!
												}
												else
												{
													$LogicalInterfaceObject += ,('IP Assign Method', $LogicalAssignmentMethod, '', '')
												}
												if ($LogicalAssignmentMethod -eq 'Static')
												{
													$LogicalInterfaceObject += ,('Primary Address', $LogicalInterface.IE.ifIPv4AddressPrimary , '', '')
													$LogicalInterfaceObject += ,('Primary Netmask', $LogicalInterface.IE.ifIPv4AddressPrimaryMask, '', '')
													if ($LogicalInterface.IE.ifIPv4ConfigSecondaryEnabled -eq '1')
													{
														$LogicalInterfaceObject += ,('Configure Secondary Interface', 'Enabled', '', '')
														$LogicalInterfaceObject += ,('Secondary Address', $LogicalInterface.IE.ifIPv4AddressSecondary, '', '')
														$LogicalInterfaceObject += ,('Secondary Netmask', $LogicalInterface.IE.ifIPv4AddressSecondaryMask, '', '')
													}
													else
													{
														if ($SWeLite)
														{
															# Don't even show the Secondary Interface reference as 'n/a this rls', as it's n/a this platform!
															if ($FormattedIfName -ne 'Admin IP')
															{
																#But Next hop only shows on the SWe Lite's Media NICs
																$LogicalInterfaceObject += ,('Media Next Hop IP', $LogicalInterface.IE.ifIPv4NextHop, '', '')
															}
														}
														else
														{
															$LogicalInterfaceObject += ,('Configure Secondary Interface', 'Disabled', '', '')
														}
													}
												}
												else
												{
													# IP v4 with DHCP.
													$LogicalInterfaceObject += ,('Primary Address', 'None' , '', '')
													$LogicalInterfaceObject += ,('Primary Netmask', 'None' , '', '')
													if ($SWeLite)
													{
														if ($FormattedIfName -ne 'Admin IP')
														{
															#Next hop only shows on the SWe Lite's Media NICs
															$LogicalInterfaceObject += ,('Media Next Hop IP', $LogicalInterface.IE.ifIPv4NextHop, '', '')
														}
													}
													$LogicalInterfaceObject += ,('DHCP Options to Use', (Test-ForNull -LookupTable $DHCPOptionsToUseLookup -value $LogicalInterface.IE.DHCPSuppliedParamUsage), '', '')
													if ($LogicalInterface.IE.DHCPSuppliedParamUsage -eq 2)
													{
														$LogicalInterfaceObject += ,('Use as Split DNS', $EnabledLookup.Get_Item($LogicalInterface.IE.ConfigureSplitDNS), '', '')
													}
												}
											}
											if ($LogicalInterfaceIpAddressingMode -ne '0')
											{
												#IP v6 or 'both'
												$LogicalInterfaceObject += ,('SPAN-L', 'IPv6 Information', '', '')
												if ($FormattedIfName -ne 'Admin IP')
												{
													$LogicalInterfaceObject += ,('ACL In', (Test-ForNull -LookupTable $ACLv6TableLookup -value $LogicalInterface.IE.ipv6AclInInstanceID), '', '')
													$LogicalInterfaceObject += ,('ACL Out', (Test-ForNull -LookupTable $ACLv6TableLookup -value $LogicalInterface.IE.ipv6AclOutInstanceID), '', '')
													$LogicalInterfaceObject += ,('ACL Forward', (Test-ForNull -LookupTable $ACLv6TableLookup -value $LogicalInterface.IE.ipv6AclForwardInstanceID), '', '')
												}
												$LogicalInterfaceObject += ,('Primary Address', $LogicalInterface.IE.ifIPv6AddressPrimary, '', '')
												$LogicalInterfaceObject += ,('Primary Address Prefix', $LogicalInterface.IE.ifIPv6AddressPrimaryPrefix, '', '')
											}
											# ---- Here we build the display names of the NICs and their IP addresses
											if ($LogicalInterface.IE.ifIPv4AddressPrimary -ne '')
											{
												$PortToIPAddressLookup.Set_Item(('{0}-1' -f $LogicalInterface.IE.IfName), ($FormattedIfName + ' (' + $LogicalInterface.IE.ifIPv4AddressPrimary + ')'))
												$PortToIPAddressLookup.Set_Item(('{0}' -f $LogicalInterface.IE.IfName), ($FormattedIfName + ' (' + $LogicalInterface.IE.ifIPv4AddressPrimary + ')')) # The SWe Lite Port version
											}
											else
											{
												$PortToIPAddressLookup.Set_Item(('{0}-1' -f $LogicalInterface.IE.IfName), '<n/a>')
											}
											if (($LogicalInterface.IE.ifIPv4ConfigSecondaryEnabled -eq 1) -and ($LogicalInterface.IE.ifIPv4AddressSecondary -ne ''))
											{
												$PortToIPAddressLookup.Set_Item(('{0}-2' -f $LogicalInterface.IE.IfName), ($FormattedIfName + ' (2nd) (' + $LogicalInterface.IE.ifIPv4AddressSecondary + ')'))
											}
											else
											{
												$PortToIPAddressLookup.Set_Item(('{0}-2' -f $LogicalInterface.IE.IfName), '<n/a>')
											}
											if ($LogicalInterfaceIpAddressingMode -ne '0')
											{
												#IP v6 or 'both'
												if ($LogicalInterface.IE.ifIPv6AddressPrimary -ne '')
												{
													$PortToIPAddressLookup.Add(('{0}-6' -f $LogicalInterface.IE.IfName), ($FormattedIfName + ' (' + $LogicalInterface.IE.ifIPv6AddressPrimary + ')'))
												}
												else
												{
													$PortToIPAddressLookup.Add(('{0}-6' -f $LogicalInterface.IE.IfName), '<n/a>')
												}
											}
											# ---- End of the NIC display naming code
										}
										$NodeHardwareLogicalInterfaceData += , ('Logical Interfaces', ('Interface Name: ' + $FormattedIfName), '', $LogicalInterfaceObject)
									}
								}
							}

							# ---- FXS Cards ----------
							if ($NodeHardwareGroup.name -eq 'FXSCards')
							{
								if ($LicencedFXSPorts -eq $null) { [int]$LicencedFXSPorts = 99} # If we've not read from the NodeInfo file, assume no restriction - show it all.
								$FXSCards = $NodeHardwareGroup.GetElementsByTagName('ID')
								$FXSCardColumnTitles = @('Card', 'Port', 'Enabled', 'Port ID', 'Port Alias', 'Port Description', 'Rx Gain', 'Tx Gain', 'Port Type', 'Analog Line Profile')
								$FXSCardCollection = @()
								if ($FXSCards.Count -ne 0)
								{
									$FxsLongLoopCardList = '' #No need for fancy arrays or lookup tables: a simple string will work here nicely
									ForEach ($FXSCard in $FXSCards)
									{
										if ($FXSCard.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($FXSCard.IE.classname -eq 'LINE_CARD_CONFIG_IE')
										{
											$FXSCardPorts = $FXSCard.GetElementsByTagName('ID')
											ForEach ($FXSCardPort in $FXSCardPorts)
											{
												if  ($FXSCardPort.IE.classname -eq 'FXS_PORT_PROFILE_CONFIG_ENTRY_IE')
												{
													if ($FXSCardPort.IE.LocalLoopType -ne $null)
													{
														#Hide even-numbered card in SBC1000 if (card-1) is set to Long Loop. (NB, this list is only populated for the SBC1k platform):
														if ($FxsLongLoopCardList -match ($FXSCard.value - 1)) { continue }
														#Now set the appropriate text and add the card to the FxsLongLoopCardList is appropriate
														if ($FXSCardPort.IE.LocalLoopType -eq '0')
														{
															$FXSPortType = 'FXS'
														}
														else
														{
															if ($platform -eq 'SBC 1000') {$FxsLongLoopCardList += ($FXSCard.value)}
															$FXSPortType = 'FXS Long Loop'
														}
													}
													else
													{
														$FXSPortType = '<n/a>' # Changed from mt usual '<n/a this rls>' to prevent the row spilling to multiple lines
													}
													$FXSCardPortName = [regex]::replace($FXSCardPort.IE.PortName, '-' , ' ') # Turn the dash into a space
													$FXSCardObject = @($FXSCard.value, $FXSCardPort.value, $EnabledLookup.Get_Item($FXSCardPort.IE.Enabled), $FXSCardPortName, $FXSCardPort.IE.PortAlias, $FXSCardPort.IE.PortDescription, ($FXSCardPort.IE.RxGain + ' dB'), ($FXSCardPort.IE.TxGain + ' dB'), $FXSPortType, $FXSFXOCountryLookup.Get_Item($FXSCardPort.IE.Country))
													if ($LicencedFXSPorts -gt 0)
													{
														$FXSCardCollection += , $FXSCardObject
														$LicencedFXSPorts --
													}
												}
											}
										}
									}
									#$NodeHardwareData += , ('FXS Cards', '', $FXSCardColumnTitles, $FXSCardCollection) - this is now added just prior to writing to Word
								}
							}
							# ---- FXO Line Cards ----------
							if ($NodeHardwareGroup.name -eq 'FXOLineCards')
							{
								if ($LicencedFXOPorts -eq $null) { [int]$LicencedFXOPorts = 99} # If we've not read from the NodeInfo file, assume no restriction - show it all.
								$FXOLineCards = $NodeHardwareGroup.GetElementsByTagName('ID')
								$FXOLineCardColumnTitles = @('Card', 'Port', 'Enabled', 'Port ID', 'Port Alias', 'Port Description', 'Rx Gain', 'Tx Gain', 'Analog Line Profile', 'Ring Validation', 'Max Freq', 'Min Freq', 'Ring Det Thresh')
								$FXOLineCardCollection = @()
								if ($FXOLineCards.Count -ne 0)
								{
									ForEach ($FXOLineCard in $FXOLineCards)
									{
										if ($FXOLineCard.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($FXOLineCard.IE.classname -eq 'LINE_CARD_CONFIG_IE')
										{
											$FXOLineCardPorts = $FXOLineCard.GetElementsByTagName('ID')
											ForEach ($FXOLineCardPort in $FXOLineCardPorts)
											{
												if  ($FXOLineCardPort.IE.classname -eq 'FXO_PORT_PROFILE_CONFIG_ENTRY_IE')
												{
													$FXOLineCardPortName = [regex]::replace($FXOLineCardPort.IE.PortName, '-' , ' ') # Turn the dash into a space
													$FXOLineCardObject = @($FXOLineCard.value, $FXOLineCardPort.value, $EnabledLookup.Get_Item($FXOLineCardPort.IE.Enabled), $FXOLineCardPortName, $FXOLineCardPort.IE.PortAlias, $FXOLineCardPort.IE.PortDescription, ($FXOLineCardPort.IE.RxGain + ' dB'), ($FXOLineCardPort.IE.TxGain + ' dB'), $FXSFXOCountryLookup.Get_Item($FXOLineCardPort.IE.Country), $EnabledLookup.Get_Item($FXOLineCardPort.IE.RingValidation), ($FXOLineCardPort.IE.MaxRingFrequency + ' Hz'), ($FXOLineCardPort.IE.MinRingFrequency + ' Hz'), ($FXORingDetectLookup.Get_Item($FXOLineCardPort.IE.RingDetectionThreshold) + ' Vrms'))
													if ($LicencedFXOPorts -gt 0)
													{
														$FXOLineCardCollection += , $FXOLineCardObject
														$LicencedFXOPorts --
													}
												}
											}
										}
									}
									#$NodeHardwareData += , ('FXO Line Cards', '', $FXOLineCardColumnTitles, $FXOLineCardCollection) - this is now added just prior to writing to Word
								}
							}
							# ---- BRI Line Cards ----------
							if ($NodeHardwareGroup.name -eq 'BRILineCards')
							{
								if ($LicencedBRIPorts -eq $null) { [int]$LicencedBRIPorts = 99} # If we've not read from the NodeInfo file, assume no restriction - show it all.
								$BRILineCards = $NodeHardwareGroup.GetElementsByTagName('ID')
								$BRILineCardColumnTitles = @('Card', 'Port', 'Enabled', 'Port ID', 'Port Alias', 'Port Description', 'Termination')
								$BRILineCardCollection = @()
								if ($BRILineCards.Count -ne 0)
								{
									ForEach ($BRILineCard in $BRILineCards)
									{
										if ($BRILineCard.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($BRILineCard.IE.classname -eq 'LINE_CARD_CONFIG_IE')
										{
											$BRILineCardPorts = $BRILineCard.GetElementsByTagName('ID')
											ForEach ($BRILineCardPort in $BRILineCardPorts)
											{
												if  ($BRILineCardPort.IE.classname -eq 'BRI_PORT_PROFILE_CONFIG_ENTRY_IE')
												{
													$BRILineCardPortName = [regex]::replace($BRILineCardPort.IE.PortName, '-' , ' ') # Turn the dash into a space
													$BRILineCardObject = @($BRILineCard.value, $BRILineCardPort.value, $EnabledLookup.Get_Item($BRILineCardPort.IE.Enabled), $BRILineCardPortName, $BRILineCardPort.IE.PortAlias, $BRILineCardPort.IE.PortDescription, $EnabledLookup.Get_Item($BRILineCardPort.IE.PortTermination))
													if ($LicencedBRIPorts -gt 0)
													{
														$BRILineCardCollection += , $BRILineCardObject
														$LicencedBRIPorts --
													}
												}
											}
										}
									}
									#$NodeHardwareData += , ('BRI Line Cards', '', $BRILineCardColumnTitles, $BRILineCardCollection) - this is now added just prior to writing to Word
								}
							}
						}
					}

					#---------- IP Routing ------------------------
					'IPRouting'
					{
						$IPRouting = $node.GetElementsByTagName('Token')
						#$IPRouteData = @() - moved July2016 to the end so we can re-order for the correct on-screen layout
						$PortMirrorData = @() #We initialise Logging Data here as we extract Port Mirroring in this section.
						ForEach ($IProute in $IPRouting)
						{
							# ---- StaticARP ----------
							if ($IPRoute.name -eq 'StaticARP')
							{
								write-verbose -message 'Skipped: Static ARP (not implemented)'
							}
							# ---- Static Routes ----------
							if ($IPRoute.name -eq 'StaticRoutes')
							{
								$StaticRouteCollection = @() #null the collection for each table
								$StaticRouteData = @()
								$StaticRoutes = $IPRoute.GetElementsByTagName('ID')
								$StaticRouteColumnTitles = @('Entry', 'IP Destination Addr', 'IP Destination Mask', 'IP Next Hop Address', 'Metric')
								if ($StaticRoutes.Count -ne 0)
								{
									ForEach ($StaticRoute in $StaticRoutes)
									{
										if ($StaticRoute.IE.classname -eq $null) { continue } # Empty / deleted entry
										$StaticRouteObject = @($StaticRoute.value, $StaticRoute.IE.ipDestinationAddr, $StaticRoute.IE.ipDestinationMask, $StaticRoute.IE.ipNexthopAddress, $StaticRoute.IE.ipStaticRouteMetric)
										$StaticRouteCollection += ,$StaticRouteObject
									}
									$StaticRouteData += ,('Static Routes', '', $StaticRouteColumnTitles, $StaticRouteCollection)
								}
							}
							# ---- PortMirrorConfig ----------
							if ($IPRoute.name -eq 'PortMirrorConfig')
							{
								$PortMirrorCollection = @() #null the collection for each table
								$PortMirrors = $IPRoute.GetElementsByTagName('ID')
								$PortMirrorColumnTitles = @('Entry', 'Mirrored Port', 'Analyzer Port', 'Direction')
								if ($PortMirrors.Count -ne 0)
								{
									ForEach ($PortMirror in $PortMirrors)
									{
										if ($PortMirror.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($PortMirror.IE.classname -eq 'NETSERV_PORT_MIRROR_CFG_IE')
										{
											$PortMirrorObject = @($PortMirror.value, $PortMirrorPortLookup.Get_Item($PortMirror.IE.PMSrcPort), $PortMirrorPortLookup.Get_Item($PortMirror.IE.PMDestPort), $PortMirrorDirectionLookup.Get_Item($PortMirror.IE.Direction))
											$PortMirrorCollection += ,$PortMirrorObject
										}
									}
									#Yes, this is being saved to LOGGING data, as it's presented to the user there.
									$PortMirrorData += ,('Port Mirror', '', $PortMirrorColumnTitles, $PortMirrorCollection)
								}
							}
							#---------- IPv4 Access Control Lists ------------
							if ($IPRoute.name -eq 'AccessCntrlList')
							{
								$accesslists = $IPRoute.GetElementsByTagName('ID')
								$AllACLData = @()
								if ($SWeLite)
								{
									$ACLColumnTitles = @('Description', 'Protocol', 'Source IP/Mask', 'Destination IP/Mask', 'Protocol Service', 'Action', 'Interface Name', 'Precedence', 'Entry')
								}
								else
								{
									$ACLColumnTitles = @('Description', 'Protocol', 'Source IP/Mask', 'Destination IP/Mask', 'Protocol Service', 'Action', 'Entry')
								}
								if ($accesslists.Count -ne 0)
								{
									ForEach ($accesslist in $accesslists)
									{
										if ($accesslist.IE.classname -eq $null) {continue }
										if ($accesslist.IE.classname -eq 'ACCESS_CONTROL_LIST_IE')
										{
											$ACLCollection = @() #null the collection for each table
											$ACLTableList = @('') 	#initialise the indexed table list
											$ACLTableHeading = Fix-NullDescription -TableDescription $accesslist.IE.Description -TableValue $accesslist.value -TablePrefix 'ACL #'
											$ACLs = $accesslist.GetElementsByTagName('ID')
											ForEach ($ACL in $ACLs)
											{
												$ACLTable = @() 	#initialise the table list
												if ($ACL.IE.classname -eq $null) { continue } # Empty / deleted entry
												if ($ACL.IE.classname -eq 'ACCESS_CONTROL_RULE_IE')
												{
													#We run through each table twice. The first time generates the overview (as a H-table) where the sequence is highlighted.
													#The second pass then pulls and builds a separate V-table for each of the individual entries
													$ACLProtocolService = $null
													if ($SWeLite)
													{
														#If both are the same, or if ONE is zero, lookup only the *Destination* port:
														if (($ACL.IE.aclSrcPort -eq $ACL.IE.aclDstPort) -or (($ACL.IE.aclSrcPort -eq 0) -or ($ACL.IE.aclDstPort -eq 0)))
														{
															$ACLProtocolService = $ACLProtocolServiceLookup.Get_Item($ACL.IE.aclDstPort)
														}
														# But the protocol might not have been on the known list, in which case the Lookup returned ''
														if ($ACLProtocolService -eq $null)
														{
															$ACLProtocolService = '--None--'
														}
													}
													else
													{
														if (($ACL.IE.aclMinDstPort -eq $ACL.IE.aclMaxDstPort) -and ($ACLProtocolServiceLookup.Get_Item($ACL.IE.aclMinDstPort) -ne $null))
														{
															$ACLProtocolService = $ACLProtocolServiceLookup.Get_Item($ACL.IE.aclMinDstPort)
														}
														else
														{
															$ACLProtocolService = '--None--'
														}
													}
													#FIRST PASS - this generates the summary H-table
													if (($ACL.IE.srcIPAddr -eq '0.0.0.0') -and ($ACL.IE.srcIPAddrMask -eq '0.0.0.0'))
													{
														$SourceIPMask = 'Any'
													}
													else
													{
														$SourceIPMask = $ACL.IE.srcIPAddr + '/' + (Convert-MaskToCIDR -Mask ($ACL.IE.srcIPAddrMask))
													}
													if (($ACL.IE.destIPAddr -eq '0.0.0.0') -and ($ACL.IE.destIPAddrMask -eq '0.0.0.0'))
													{
														$DestIPMask = 'Any'
													}
													else
													{
														$DestIPMask = $ACL.IE.destIPAddr + '/' + (Convert-MaskToCIDR -Mask ($ACL.IE.destIPAddrMask))
													}
													$ACLProtocol = $ACLProtocolLookup.Get_Item($ACL.IE.aclProtocol)
													if ($ACLProtocol -eq $null) { $ACLProtocol = $ACL.IE.aclProtocol } #If the number isn't in the list, just report its raw number
													if ($SWeLite)
													{
														$ACLObject = @($ACL.IE.Description, $ACLProtocol, $SourceIPMask, $DestIPMask, $ACLProtocolService, $ACLAllowLookup.Get_Item($ACL.IE.aclAction), $ACL.IE.ifName, $ACL.IE.aclPrecedence, $ACL.value)
													}
													else
													{
														$ACLObject = @($ACL.IE.Description, $ACLProtocol, $SourceIPMask, $DestIPMask, $ACLProtocolService, $ACLAllowLookup.Get_Item($ACL.IE.aclAction), $ACL.value)
													}
													$ACLCollection += ,$ACLObject
													#SECOND PASS - this generates a V-table for each entry in the table
													$ACLTableEntryTitle = 'Entry ' + $ACL.Value + ' : ' + $ACL.IE.Description
													$ACLTable += ,('SPAN', $ACLTableEntryTitle, '', '')
													$ACLTable += ,('Description', $ACL.IE.Description, '', '')
													if ($ACLProtocolLookup.Get_Item($ACL.IE.aclProtocol) -eq $null)
													{
														$ACLTable += ,('Protocol', 'Other', '', '')
														$ACLTable += ,('IANA IP Protocol Number ', $ACLProtocol, '', '')
													}
													else
													{
														$ACLTable += ,('Protocol', $ACLProtocol, '', '')
													}
													$ACLTable += ,('Action', $ACLAllowLookup.Get_Item($ACL.IE.aclAction), '', '')
													if (($ACL.IE.aclProtocol -eq '6') -or ($ACL.IE.aclProtocol -eq '17')) # (TCP or UDP)
													{
														if ($ACLProtocolService -ne '--None--')
														{
															$ACLTable += ,('Port Selection Method', 'Service', '', '')
															$ACLTable += ,('Service', $ACLProtocolService, '', '')
														}
														else
														{
															if ($SWeLite)
															{
																$ACLTable += ,('Port Selection Method', 'Single Port', '', '')
															}
															else
															{
																$ACLTable += ,('Port Selection Method', 'Range', '', '')
															}
														}
													}
													if ($SWeLite)
													{
														$ACLTable += ,('Precedence', ($ACL.IE.aclPrecedence + ' [1..65535]'), '', '')
														$ACLTable += ,('Bucket Size', ($ACL.IE.aclBucketSize + ' [0..255]'), '', '')
														$ACLTable += ,('Fill Rate', ($ACL.IE.aclFillRate + ' [0..25000]'), '', '')
														if ($ACL.IE.ifName -eq '')
														{
															$ACLTable += ,('Interface Name', 'Any', '', '')
														}
														else
														{
															$ACLTable += ,('Interface Name', $PortToIPAddressLookup.Get_Item($ACL.IE.ifName), '', '')
														}
													}
													$ACLTable += ,('SPAN-L', 'Source', 'SPAN-R', 'Destination')
													$ACLTable += ,('IP Address', $ACL.IE.srcIPAddr, 'IP Address', $ACL.IE.destIPAddr)
													$ACLTable += ,('Netmask', $ACL.IE.srcIPAddrMask, 'Netmask', $ACL.IE.destIPAddrMask)
													if (($ACL.IE.aclProtocol -eq '6') -or ($ACL.IE.aclProtocol -eq '17')) # (TCP or UDP)
													{
														if ($ACLProtocolService -eq '--None--')
														{
															if ($SWeLite)
															{
																$ACLTable += ,('Port Number', $ACL.IE.aclSrcPort , 'Minimum Port Number', $ACL.IE.aclDstPort)
															}
															else
															{
																$ACLTable += ,('Minimum Port Number', $ACL.IE.aclMinSrcPort , 'Minimum Port Number', $ACL.IE.aclMinDstPort)
																$ACLTable += ,('Maximum Port Number', $ACL.IE.aclMaxSrcPort, 'Maximum Port Number', $ACL.IE.aclMaxDstPort)
															}
														}
													}
													#Add sufficient empty values into the array so that we can then poke the value in at the appropriate level
													# (there might be only 3 values, but they might be entries 1, 2 & 5 after the user has deleted some)
													while (($ACLTableList.Count -1) -lt $ACL.Value) { $ACLTableList += '' }
													$ACLTableList[$ACL.Value] = $ACLTable #Stick each table into an indexed array so we can re-order below & display in the correct sequence
													$ACLTable = @() #Initialise the table for next time. This is needed here in case an empty table follows - it ensures the below code realises and handles correctly
												}
											}
											#Now we need to rearrange the entries in the table into their correct sequence:
											$OrderedACLTableList = @()
											if (($accesslist.IE.Sequence -ne '') -and ($accesslist.IE.Sequence -ne $null)) #If it exists AND it's not blank
											{
												$sequence = Decode-Sequence -EncodedSequence $accesslist.IE.Sequence -byteCount 2
												$AclTemp = @()
												foreach ($reindex in $sequence)
												{
													foreach ($tableRow in $ACLCollection)
													{
														if ($SWeLite)
														{
															$IndexVar = 8
														}
														else
														{
															$IndexVar = 6
														}
														if ($tableRow[$IndexVar] -eq $reindex)
														{
															$AclTemp += ,$TableRow
														}
													}
													$OrderedACLTableList += ,($ACLTableList[$reindex])
												}
												$ACLCollection = $AclTemp
											}
											else
											{
												#No sequence? Then we only have one value in the $ACLTableList array - which also equals $ACLTable.
												$OrderedACLTableList = $ACLTableList
											}
											#This saves the overview/heading table:
											$AllACLData += ,('Access Control Lists', $ACLTableHeading, $ACLColumnTitles, $ACLCollection)
											#This loop saves the underlying tables (if there are any):
											foreach ($OneTable in $OrderedACLTableList)
											{
												if ($OneTable -eq '')
												{
													# It's empty. (This will be the null entry we initialised the array with
													continue
												}
												$AllACLData += ,('Access Control Lists', '', '', $OneTable)
											}
										}
									}
								}
							}
							#---------- IPv6 Access Control Lists ------------
							if ($IPRoute.name -eq 'IPv6AccessCntrlList')
							{
								$accesslists = $IPRoute.GetElementsByTagName('ID')
								$AllACLv6Data = @()
								if ($SWeLite)
								{
									$ACLv6ColumnTitles = @('Description', 'Protocol', 'Source IP/Mask', 'Destination IP/Mask', 'Protocol Service', 'Action', 'Interface Name', 'Precedence', 'Entry')
								}
								else
								{
									$ACLv6ColumnTitles = @('Description', 'Protocol', 'Source IP/Mask', 'Destination IP/Mask', 'Protocol Service', 'Action', 'Entry')
								}
								if ($accesslists.Count -ne 0)
								{
									ForEach ($accesslist in $accesslists)
									{
										if ($accesslist.IE.classname -eq $null) {continue }
										if ($accesslist.IE.classname -eq 'ACCESS_CONTROL_LIST_IE')
										{
											$ACLv6Collection = @() #null the collection for each table
											$ACLv6TableList = @('') 	#initialise the indexed table list
											$ACLv6TableHeading = Fix-NullDescription -TableDescription $accesslist.IE.Description -TableValue $accesslist.value -TablePrefix 'ACL #'
											$v6ACLs = $accesslist.GetElementsByTagName('ID')
											ForEach ($v6ACL in $v6ACLs)
											{
												$ACLv6Table = @() 	#initialise the table list
												if ($v6ACL.IE.classname -eq $null) { continue } # Empty / deleted entry
												if ($v6ACL.IE.classname -eq 'ACCESS_CONTROL_RULE_IE')
												{
													#We run through each table twice. The first time generates the overview (as a H-table) where the sequence is highlighted.
													#The second pass then pulls and builds a separate V-table for each of the individual entries
													$ACLv6ProtocolService = $null
													if (($v6ACL.IE.aclDstPort -ne $null) -and ($ACLProtocolServiceLookup.Get_Item($v6ACL.IE.aclDstPort) -ne $null))
													{
														#Specific to the SweLite at the moment (v7.0)
														$ACLv6ProtocolService = $ACLProtocolServiceLookup.Get_Item($v6ACL.IE.aclDstPort)
													}
													elseif (($v6ACL.IE.aclMinDstPort -eq $v6ACL.IE.aclMaxDstPort) -and ($ACLProtocolServiceLookup.Get_Item($v6ACL.IE.aclMinDstPort) -ne $null))
													{
														$ACLv6ProtocolService = $ACLProtocolServiceLookup.Get_Item($v6ACL.IE.aclMinDstPort)
													}
													else
													{
														$ACLv6ProtocolService = '--None--'
													}
													#FIRST PASS - this generates the summary H-table
													if ($v6ACL.IE.srcIPAddr -eq '::')
													{
														$SourceIPMask = 'Any'
													}
													else
													{
														$SourceIPMask = $v6ACL.IE.srcIPAddr + '/' + $v6ACL.IE.srcIPAddrMask
													}
													if ($v6ACL.IE.destIPAddr -eq '::')
													{
														$DestIPMask = 'Any'
													}
													else
													{
														$DestIPMask = $v6ACL.IE.destIPAddr + '/' + $v6ACL.IE.destIPAddrMask
													}
													$v6ACLProtocol = $ACLProtocolLookup.Get_Item($v6ACL.IE.aclProtocol)
													if ($v6ACLProtocol -eq $null) { $v6ACLProtocol = $v6ACL.IE.aclProtocol } #If the number isn't in the list, just report its raw number
													if ($SWeLite)
													{
														$ACLv6Object = @($v6ACL.IE.Description, $v6ACLProtocol, $SourceIPMask, $DestIPMask, $ACLv6ProtocolService, $ACLAllowLookup.Get_Item($v6ACL.IE.aclAction), $v6ACL.IE.ifName, $v6ACL.IE.aclPrecedence, $v6ACL.value)
													}
													else
													{
														$ACLv6Object = @($v6ACL.IE.Description, $v6ACLProtocol, $SourceIPMask, $DestIPMask, $ACLv6ProtocolService, $ACLAllowLookup.Get_Item($v6ACL.IE.aclAction), $v6ACL.value)
													}
													$ACLv6Collection += ,$ACLv6Object
													#SECOND PASS - this generates a V-table for each entry in the table
													$ACLv6TableEntryTitle = 'Entry ' + $v6ACL.Value + ' : ' + $v6ACL.IE.Description
													$ACLv6Table += ,('SPAN', $ACLv6TableEntryTitle, '', '')
													$ACLv6Table += ,('Description', $v6ACL.IE.Description, '', '')
													if ($ACLProtocolLookup.Get_Item($v6ACL.IE.aclProtocol) -eq $null)
													{
														$ACLv6Table += ,('Protocol', 'Other', '', '')
														$ACLv6Table += ,('IANA Protocol Number ', $v6ACLProtocol, '', '')
													}
													else
													{
														$ACLv6Table += ,('Protocol', $v6ACLProtocol, '', '')
													}
													$ACLv6Table += ,('Action', $ACLAllowLookup.Get_Item($v6ACL.IE.aclAction), '', '')
													if (($v6ACL.IE.aclProtocol -eq '6') -or ($v6ACL.IE.aclProtocol -eq '17')) # (TCP or UDP)
													{
														if ($ACLv6ProtocolService -ne '--None--')
														{
															$ACLv6Table += ,('Port Selection Method', 'Service', '', '')
															$ACLv6Table += ,('Service', $ACLv6ProtocolService, '', '')
														}
														else
														{
															if ($SWeLite)
															{
																$ACLv6Table += ,('Port Selection Method', 'Single Port', '', '')
															}
															else
															{
																$ACLv6Table += ,('Port Selection Method', 'Range', '', '')
															}
														}
													}
													if ($SWeLite)
													{
														$ACLv6Table += ,('Precedence', ($v6ACL.IE.aclPrecedence + ' [1..65535]'), '', '')
														$ACLv6Table += ,('Bucket Size', ($v6ACL.IE.aclBucketSize + ' [0..255]'), '', '')
														$ACLv6Table += ,('Fill Rate', ($v6ACL.IE.aclFillRate + ' [0..25000]'), '', '')
														if ($v6ACL.IE.ifName -eq '')
														{
															$ACLv6Table += ,('Interface Name', 'Any', '', '')
														}
														else
														{
															$ACLv6Table += ,('Interface Name', $PortToIPAddressLookup.Get_Item($v6ACL.IE.ifName), '', '')
														}
													}
													$ACLv6Table += ,('SPAN-L', 'Source', 'SPAN-R', 'Destination')
													$ACLv6Table += ,('IP Address', $v6ACL.IE.srcIPAddr, 'IP Address', $v6ACL.IE.destIPAddr)
													$ACLv6Table += ,('Network Prefix', $v6ACL.IE.srcIPAddrMask, 'Network Prefix', $v6ACL.IE.destIPAddrMask)
													if (($v6ACL.IE.aclProtocol -eq '6') -or ($v6ACL.IE.aclProtocol -eq '17')) # (TCP or UDP)
													{
														if ($ACLv6ProtocolService -eq '--None--')
														{
															if ($SWeLite)
															{
																$ACLv6Table += ,('Port Number', $v6ACL.IE.aclSrcPort , 'Port Number', $v6ACL.IE.aclDstPort)
															}
															else
															{
																$ACLv6Table += ,('Minimum Port Number', $v6ACL.IE.aclMinSrcPort , 'Minimum Port Number', $v6ACL.IE.aclMinDstPort)
																$ACLv6Table += ,('Maximum Port Number', $v6ACL.IE.aclMaxSrcPort, 'Maximum Port Number', $v6ACL.IE.aclMaxDstPort)
															}
														}
													}
													#Add sufficient empty values into the array so that we can then poke the value in at the appropriate level
													# (there might be only 3 values, but they might be entries 1, 2 & 5 after the user has deleted some)
													while (($ACLv6TableList.Count -1) -lt $v6ACL.Value) { $ACLv6TableList += '' }
													$ACLv6TableList[$v6ACL.Value] = $ACLv6Table #Stick each table into an indexed array so we can re-order below & display in the correct sequence
													$ACLv6Table = @() #Initialise the table for next time. This is needed here in case an empty table follows - it ensures the below code realises and handles correctly
												}
											}
											#Now we need to rearrange the entries in the table into their correct sequence:
											$OrderedACLv6TableList = @()
											if (($accesslist.IE.Sequence -ne '') -and ($accesslist.IE.Sequence -ne $null)) #If it exists AND it's not blank
											{
												$sequence = Decode-Sequence -EncodedSequence $accesslist.IE.Sequence -byteCount 2
												$ACLv6Temp = @()
												foreach ($reindex in $sequence)
												{
													foreach ($tableRow in $ACLv6Collection)
													{
														if ($SWeLite)
														{
															$IndexVar = 8
														}
														else
														{
															$IndexVar = 6
														}
														if ($tableRow[$IndexVar] -eq $reindex)
														{
															$ACLv6Temp += ,$TableRow
														}
													}
													$OrderedACLv6TableList += ,($ACLv6TableList[$reindex])
												}
												$ACLv6Collection = $ACLv6Temp
											}
											else
											{
												#No sequence? Then we only have one value in the $ACLTableList array - which also equals $ACLTable.
												$OrderedACLv6TableList = $ACLv6TableList
											}
											#This saves the overview/heading table:
											$AllACLv6Data += ,('IPv6 Access Control Lists', $ACLv6TableHeading, $ACLv6ColumnTitles, $ACLv6Collection)
											#This loop saves the underlying tables (if there are any):
											foreach ($OneTable in $OrderedACLv6TableList)
											{
												if ($OneTable -eq '')
												{
													# It's empty. (This will be the null entry we initialised the array with
													continue
												}
												$AllACLv6Data += ,('IPv6 Access Control Lists', '', '', $OneTable)
											}
										}
									}
								}
							}
							# ---- Network Monitoring ----------
							if ($IPRoute.name -eq 'LinkMonitorConfig')
							{
								$NetworkMonitoringData = @()
								$LinkMonitors = $IPRoute.GetElementsByTagName('ID')
								if ($LinkMonitors.Count -ne 0)
								{
									ForEach ($MonitoredHost in $LinkMonitors)
									{
										if ($MonitoredHost.IE.classname -eq 'NETSERV_LINK_MONITOR_CFG_IE')
										{
											$LinkMonitorTable = @() #null the collection for each table
											$LinkMonitorTable += ,('SPAN-L', 'Link Monitor Table', '', '')
											$LinkMonitorDescription = (Fix-NullDescription -TableDescription $MonitoredHost.IE.GWDescription -TableValue $MonitoredHost.value -TablePrefix 'Link Monitor #')
											$LinkMonitorTable += ,('Description', $LinkMonitorDescription, '' , '')
											$LinkMonitorTable += ,('Service Status', '<Not captured>', '' , '')
											$LinkMonitorTable += ,('SPAN-L', 'Settings', '', '')
											switch ($MonitoredHost.IE.HostType)
											{
												'0'
												{
													# Host
													$LinkMonitorTable += ,('Monitor Type', ($LinkMonitorTypeLookup.Get_Item($MonitoredHost.IE.MonitorType)), '' , '')
													$LinkMonitorTable += ,('Monitored Peer Type', 'Host', '' , '')
													$LinkMonitorTable += ,('Host Name', $MonitoredHost.IE.HostIpAddress, '' , '')
													$LinkMonitorTable += ,('Host IP Address', '<Not captured>', '' , '')
													$LinkMonitorTable += ,('Poll Period', ($MonitoredHost.IE.TimeProbePeriod + ' secs [5..60]'), '' , '')
													$LinkMonitorTable += ,('Failed Polls for Down State', $MonitoredHost.IE.PollToDownState, '' , '')
													$LinkMonitorTable += ,('Successful Polls for Up State', $MonitoredHost.IE.PollToServiceReadyState, '' , '')
												}
												'1'
												{
													# Gateway
													$LinkMonitorTable += ,('Monitor Type', ($LinkMonitorTypeLookup.Get_Item($MonitoredHost.IE.MonitorType)), '' , '')
													$LinkMonitorTable += ,('Monitored Peer Type', 'Gateway', '' , '')
													$LinkMonitorTable += ,('Gateway IP Address', $MonitoredHost.IE.HostIpAddress, '' , '')
													$LinkMonitorTable += ,('Lync CAC Profile', ($CACProfileLookup.Get_Item($MonitoredHost.IE.CACProfileID)), '' , '')
													$LinkMonitorTable += ,('Poll Period', ($MonitoredHost.IE.TimeProbePeriod + ' secs [5..60]'), '' , '')
													$LinkMonitorTable += ,('Failed Polls for Down State', $MonitoredHost.IE.PollToDownState, '' , '')
													$LinkMonitorTable += ,('Successful Polls for Up State', $MonitoredHost.IE.PollToServiceReadyState, '' , '')
													$IPSecTunnelList = ''
													$IPSecTunnelList += ($IPSecTunnelLookup.Get_Item($MonitoredHost.IE.IdIPsec_1) + "`n")
													$IPSecTunnelList += ($IPSecTunnelLookup.Get_Item($MonitoredHost.IE.IdIPsec_2) + "`n")
													$IPSecTunnelList += ($IPSecTunnelLookup.Get_Item($MonitoredHost.IE.IdIPsec_3) + "`n")
													$IPSecTunnelList += ($IPSecTunnelLookup.Get_Item($MonitoredHost.IE.IdIPsec_4) + "`n")
													$IPSecTunnelList = Strip-TrailingCR -DelimitedString $IPSecTunnelList
													$LinkMonitorTable += ,('Associated IPsec Tunnels', $IPSecTunnelList, '' , '')
												}
												default
												{
													$LinkMonitorTable += ,('Monitor Type', '<Unhandled Value>', '' , '')
												}
											}
											$NetworkMonitoringData += ,('Link Monitors', $LinkMonitorDescription, '', $LinkMonitorTable)
										}
									}
								}
							}
							# ---- RIPConfiguration ----------
							if ($IPRoute.name -eq 'RIPConfiguration')
							{
							   write-verbose -message 'Skipped: RIP Configuration (not implemented)'
							}
							# ---- DRInterfaceCfg ----------
							if ($IPRoute.name -eq 'DRInterfaceCfg')
							{
								write-verbose -message 'Skipped: DR Interface Configuration (not implemented)'
							}
							# ---- OSPFConfiguration ----------
							if ($IPRoute.name -eq 'OSPFConfiguration')
							{
								write-verbose -message 'Skipped: OSPF Configuration (not implemented)'
							}
							# ---- NATInterface ----------
							if ($IPRoute.name -eq 'NATInterface')
							{
								write-verbose -message 'Skipped: NAT Interface configuration (not implemented)'
							}
							# ---- NATPortForward ----------
							if ($IPRoute.name -eq 'NATPortForward')
							{
								write-verbose -message 'Skipped: NAT Port Forward configuration (not implemented)'
							}
						}
					}

					#---------- AD ------------------------
					'AD'
					{
						$ADData = @()
						$ADAttributes = [regex]::replace($node.IE.Attributes, ',' , "`n") # Write each attribute on a new line
						$ADTable = @()
						$ADTable += ,('SPAN-L', 'Active Directory Configuration', '', '')
						$ADTable += ,('AD Enabled', $TrueFalseLookup.Get_Item($node.IE.Enabled), '', '')
						$ADTable += ,('Use TLS', $TrueFalseLookup.Get_Item($node.IE.UseTLS), '', '')
						$ADTable += ,('Operating Mode', $ADOperatingModeLookup.Get_Item($node.IE.OperatingMode), '', '')
						if ($node.IE.OperatingMode -eq '2')
						{
							#Auth-Only. Nothing to cache!
						}
						else
						{
							#Everything else
							$ADTable += ,('Query/Cache Attributes', $ADAttributes, '', '')
						}
						$ADTable += ,('Nested Group Lookup for Authentication', (Test-ForNull -LookupTable $TrueFalseLookup -value $node.IE.NestedGroupLookups), '', '')
						if (($node.IE.OperatingMode -eq '0') -or ($node.IE.OperatingMode -eq '3'))
						{
							$ADTable += ,('SPAN-L', 'Cache Settings', '', '')
							$ADTable += ,('Normalize Cache', (Test-ForNull -LookupTable $TrueFalseLookup $node.IE.NormalizeCache), '', '')
							$ADTable += ,('Update Frequency', (Test-ForNull -LookupTable $null -value $node.IE.UpdateFrequency), '', '')
							if ($node.IE.InitialUpdateByTime -eq '1')
							{
								$ADTable += ,('Configure Initial Update Time', 'True', '', '')
								$ADTable += ,('First Update Time', $node.IE.FirstUpdateTime, '', '')
							}
							else
							{
								$ADTable += ,('Configure Initial Update Time', 'False', '', '')
							}
							$ADTable += ,('AD Backup Failure Alarm', (Test-ForNull -LookupTable $EnableLookup -value $node.IE.ADBackup), '', '')
							$ADTable += ,('Encrypt AD Cache', (Test-ForNull -LookupTable $TrueFalseLookup -value $node.IE.ADCacheEncrypt), '', '')
						}
						$ADData += ,('Active Directory', 'Configuration', '', $ADTable)

						$DomainControllers = $node.GetElementsByTagName('Token')
						if ($DomainControllers.name -eq 'DC')
						{
							$DCprofiles = $DomainControllers.GetElementsByTagName('ID')
							ForEach ($DCprofile in $DCprofiles)
							{
								if ($DCprofile.IE.classname -eq $null) { continue } # Empty / deleted entry
								if ($DCprofile.IE.classname -eq 'ADS_AD_DOMAINCONTROLLERS')
								{
									$DCTable = @()
									if ($DCprofile.IE.DCPriority -eq '0')
									{
										$DCprofiledcpriority = 'Unassigned'
									}
									else
									{
										$DCprofiledcpriority = Test-ForNull -LookupTable $null -value $dcprofile.ie.dcpriority
									}
									if ($DCprofile.IE.Description -eq '')
									{
										$DCDescription = $DCprofile.IE.DomainController
									}
									else
									{
										$DCDescription = $DCprofile.IE.Description
									}

									$DCServerTimeout = Test-ForNull -LookupTable $null -value $DCprofile.IE.LdapTimeout
									if ($DCServerTimeout -ne '<n/a this rls>')
									{
											$DCServerTimeout += ' secs [5...15]'
									}
									$DCUsername = Test-ForNull -LookupTable $null -value $DCprofile.IE.UserName
									$DCprofileIESearchScope = [regex]::replace($DCprofile.IE.SearchScope, ',' , ', ')
									$DCTable += ,('SPAN-L', $DCDescription, '', '')
									$DCTable += ,('DC Enabled', (Test-ForNull -LookupTable $EnabledLookup -value $DCprofile.IE.Enabled), '', '')
									$DCTable += ,('Description', $DCprofile.IE.Description, '', '')
									$DCTable += ,('Domain Controller Address', $DCprofile.IE.DomainController, '', '')
									$DCTable += ,('Preferred IP Version', (Test-ForNull -LookupTable $IpVersionLookup -value $DCprofile.IE.PreferredIpVersion), '', '')
									$DCTable += ,('DC Type', (Test-ForNull -LookupTable $DCTypeLookup -value $DCprofile.IE.DCType), '', '')
									$DCTable += ,('Search Scope', $DCprofileIESearchScope, '', '')
									if ($DCprofile.IE.DCType -eq '1') # Call Route
									{
										$DCTable += ,('LDAP Query', $DCprofile.IE.LDAPQuery, '', '')
										$DCTable += ,('Server Timeout', ($DCServerTimeout), '', '')
										$DCTable += ,('User Name', $DCUsername, '', '')
										$DCTable += ,('Password Setting', 'Use Current', '', '')
										if ($DCprofile.IE.DCRole -eq '0')
										{
											$DCTable += ,('DC Role', 'Primary', '', '')
										}
										else
										{
											$DCTable += ,('DC Role', 'Backup', '', '')
										}
									}
									else #Auth or On-Premises
									{
										$DCTable += ,('Server Timeout', ($DCServerTimeout), '', '')
										$DCTable += ,('User Name', $DCUsername, '', '')
										$DCTable += ,('Password Setting', 'Use current', '', '')
									}
									$DCTable += ,('DC Priority', $DCprofileDCPriority, '', '')
									$ADData += , ('Active Directory', ('Domain Controller: ' + $DCDescription), '', $DCTable)
								}
							}
						}
					}

					#---------- SBA ------------------------
					'SBA'
					{
						if ($node.IE.Classname -eq 'SBA_CONFIGURATION')
						{
							if ($node.IE.Enabled -eq '1')
							{
								$SBATable = @()
								$SBAData = @()
								$SBATable += ,('SPAN', 'ASM Configuration', '', '')
								$SBATable += ,('Remote Desktop Enabled', $YesNoLookup.Get_Item($node.IE.RemoteDesktopEnabled), '', '')
								$SBATable += ,('Windows Firewall Enabled', '<Not Available>', '', '')
								if ($node.IE.ProxyEnabled -eq $null)
								{
									$SBATable += ,('Proxy Enabled', '<n/a this rls>', '', '')
								}
								else
								{
									if ($node.IE.ProxyEnabled -eq '1')
									{
										$SBATable += ,('Proxy Enabled', 'Yes', '', '')
										$SBATable += ,('Proxy Address', ($node.IE.ProxyAddress + ' [1..65000]'), '', '')
										$SBATable += ,('Proxy Port', $node.IE.ProxyPort, '', '')
									}
									else
									{
										$SBATable += ,('Proxy Enabled', 'No', '', '')
									}
								}
								$SBATable += ,('SPAN', 'Network Adapter 1 Configuration', '', '')
								# 'IP Addressing Mode' goes here
								if ($node.IE.DHCPEnabled -eq '1')
								{
									$SBATable += ,('DHCP Enabled', 'Yes', '', '')
								}
								else
								{
									$SBATable += ,('DHCP Enabled', 'No', '', '')
									$SBATable += ,('SPAN-L', 'IP Address', 'SPAN-R', 'DNS Addresses')
									$SBATable += ,('IP Address', $node.IE.ipv4Address , 'Preferred DNS', $node.IE.DNSServer1)
									$SBATable += ,('Subnet Mask', $node.IE.ipv4Netmask , 'Alternate DNS', $node.IE.DNSServer2)
									$SBATable += ,('Default Gateway', $node.IE.ipv4Gateway , '', '')
								}
								if ($ASM_WindowsEthernetSecMac -eq 'Not Available')
								{
									# Maybe, maybe not - but we need the nodeinfo to know for sure
									$SBATable += ,('SPAN', 'Network Adapter 2 Configuration', '', '')
									$SBATable += ,('Not Available', '(need nodeinfo.txt)', '', '')
								}
								elseif ($ASM_WindowsEthernetSecMac -ne 'None')
								{
									$SBATable += ,('SPAN', 'Network Adapter 2 Configuration', '', '')
									# 'IP Addressing Mode' goes here
									if ($node.IE.Eth2ipv4DHCPEnabled -eq '1')
									{
										$SBATable += ,('DHCP Enabled', 'Yes', '', '')
									}
									else
									{
										$SBATable += ,('DHCP Enabled', 'No', '', '')
										$SBATable += ,('SPAN-L', 'IP Address', 'SPAN-R', 'DNS Addresses')
										$SBATable += ,('IP Address', $node.IE.Eth2ipv4Address , 'Preferred DNS', $node.IE.Eth2DNSServer1)
										$SBATable += ,('Subnet Mask', $node.IE.Eth2ipv4Netmask , 'Alternate DNS', $node.IE.Eth2DNSServer2)
										$SBATable += ,('Default Gateway', $node.IE.Eth2ipv4Gateway , '', '')
									}
								}
								$SBAData += ,('ASM Configuration', '', '', $SBATable)
							}
						}
						$SBAGroups = $node.GetElementsByTagName('Token')
						ForEach ($SBAGroup in $SBAGroups)
						{
							# ---- Skype / Lync CAC Profiles ----------
							if ($SBAGroup.name -eq 'CACProfile')
							{
								$CACProfileColumnTitles = @('Description', 'Skype / Lync Profile Name', 'Link Audio Bandwidth (kbps)', 'Session Audio Bandwidth (kbps)', 'Link Video Bandwidth (kbps)', 'Session Video Bandwidth (kbps)')
								$CACProfileCollection = @()
								$CACProfiles = $node.GetElementsByTagName('ID')
								if ($CACProfiles.Count -ne 0)
								{
									ForEach ($CACProfile in $CACProfiles)
									{
										if ($CACProfile.IE.classname -eq 'CAC_PROFILE_CONFIG')
										{
											if ($CACProfile.IE.TotalLinkAudioLimit -eq '0')
											{
												$CACProfileTotalLinkAudioLimit = 'Disabled'
												$CACProfileSessionAudioLimit = 'N/A'
											}
											else
											{
												$CACProfileTotalLinkAudioLimit = $CACProfile.IE.TotalLinkAudioLimit
												$CACProfileSessionAudioLimit = $CACProfile.IE.SessionAudioLimit
											}
											if ($CACProfile.IE.TotalLinkVideoLimit -eq '0')
											{
												$CACProfileTotalLinkVideoLimit = 'Disabled'
												$CACProfileSessionVideoLimit = 'N/A'
											}
											else
											{
												$CACProfileTotalLinkVideoLimit = $CACProfile.IE.TotalLinkVideoLimit
												$CACProfileSessionVideoLimit = $CACProfile.IE.SessionVideoLimit
											}
											$CACProfileObject = @($CACProfile.IE.Description, $CACProfile.IE.CACProfileName, $CACProfileTotalLinkAudioLimit, $CACProfileSessionAudioLimit, $CACProfileTotalLinkVideoLimit, $CACProfileSessionVideoLimit)
											$CACProfileCollection += , $CACProfileObject
										}
									}
									#These live under SBA but show on-screen under Protocols / Network Monitoring
									$CACData += ,('Network Monitoring', 'Skype / Lync CAC Profiles', $CACProfileColumnTitles, $CACProfileCollection)
								}
							}
						}
					}

					#---------- System ------------------------
					'System'
					{
						$systemgroups = $node.GetElementsByTagName('Token')
						$CertificateCollectionMy = @() # These need to be initialised up here as I'm reusing the same code block for both 'Certificates' & 'SupplementCertificates'
						$CertificateCollectionRoot = @()
						ForEach ($systemgroup in $systemgroups)
						{
							# ---- Certificates ----------
							if (($systemgroup.name -eq 'Certificates') -or ($systemgroup.name -eq 'SupplementCertificates'))
							{
								$certs = $systemgroup.GetElementsByTagName('ID')
								if ($certs.Count -ne 0)
								{
									ForEach ($cert in $certs)
									{
										if ($cert.IE.classname -eq 'CERTIFICATE_FILE_DATA_CFG_IE')
										{
											$CertCreateDate = ([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($cert.IE.CertStartDateEpoch))).ToString('MMM dd, yyyy HH:mm:ss')
											$CertExpiryDate = ([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($cert.IE.CertEndDateEpoch))).ToString('MMM dd, yyyy HH:mm:ss')
											$CertificateObject = @()
											#I initially used 'convertfrom-stringdata' here but it err'd when it encountered 'DC=' multiply in the string.
											$CertValueC,  $CertValueST,  $CertValueL,  $CertValueO,  $CertValueOU,  $CertValueE,  $CertIssuerCN = ''
											$CertIssuerC, $CertIssuerST, $CertIssuerL, $CertIssuerO, $CertIssuerOU, $CertIssuerE, $CertIssuerCN = ''
											#It's a little ugly, but this adds a CR if we match on "slash + 1 or 2 chars followed by ="
											$CertValuesCR = [regex]::replace($cert.IE.CertSubjectName,'/([a-zA-Z]{1,12}=)',"`n`$1")
											#... and this splits on the newline:
											$CertValues = ($CertValuesCR).Split([Environment]::NewLine, [System.StringSplitOptions]::RemoveEmptyEntries)
											foreach ($CertValue in $CertValues)
											{
												if ($CertValue.StartsWith('C='))  { $CertValueC  = $CertValue.Substring(2) }
												if ($CertValue.StartsWith('ST=')) { $CertValueST = $CertValue.Substring(3) }
												if ($CertValue.StartsWith('L='))  { $CertValueL  = $CertValue.Substring(2) }
												if ($CertValue.StartsWith('O='))  { $CertValueO  = $CertValue.Substring(2) }
												if ($CertValue.StartsWith('OU=')) { $CertValueOU = $CertValue.Substring(3) }
												if ($CertValue.StartsWith('E='))  { $CertValueE  = $CertValue.Substring(2) }
												if ($CertValue.StartsWith('emailAddress='))  { $CertValueE  = $CertValue.Substring(13) }
											}
											$CertIssuerCR = [regex]::replace($cert.IE.CertIssuerName,'/([a-zA-Z]{1,12}=)',"`n`$1")
											$CertIssuerValues = ($CertIssuerCR).Split([Environment]::NewLine, [System.StringSplitOptions]::RemoveEmptyEntries)
											foreach ($CertIssuerValue in $CertIssuerValues)
											{
												if ($CertIssuerValue.StartsWith('CN=')) { $CertIssuerCN = $CertIssuerValue.Substring(3) }
												if ($CertIssuerValue.StartsWith('C='))  { $CertIssuerC  = $CertIssuerValue.Substring(2) }
												if ($CertIssuerValue.StartsWith('ST=')) { $CertIssuerST = $CertIssuerValue.Substring(3) }
												if ($CertIssuerValue.StartsWith('L='))  { $CertIssuerL  = $CertIssuerValue.Substring(2) }
												if ($CertIssuerValue.StartsWith('O='))  { $CertIssuerO  = $CertIssuerValue.Substring(2) }
												if ($CertIssuerValue.StartsWith('OU=')) { $CertIssuerOU = $CertIssuerValue.Substring(3) }
												if ($CertIssuerValue.StartsWith('E='))  { $CertIssuerE  = $CertIssuerValue.Substring(2) }
												if ($CertIssuerValue.StartsWith('emailAddress=')) { $CertIssuerE  = $CertIssuerValue.Substring(13) }
											}
											$CertificateObject += ,('SPAN-L', 'Subject', 'SPAN-R', 'Issuer')
											$CertCommonName = $cert.IE.CertCommonName
											if ($CertCommonName -eq "") { $CertCommonName = "[No CN]" } #Found some GoDaddy roots have no CN!
											#OK, this repetition is a bit ugly, but splitting into LH and RH columns here seemed unnecessarily complex.
											if ($cert.IE.CertIssuerName -match $cert.IE.CertSubjectName)
											{
												# It's self-signed. The RH side of the column can be skipped
												$CertificateObject += ,('Common Name', $CertCommonName, '', 'Certificate is Self-Signed')
												$CertificateObject += ,('ISO Country Code', $CertValueC, '', '')
												$CertificateObject += ,('State or Province', $CertValueST, '', '')
												$CertificateObject += ,('Locality', $CertValueL, '', '')
												$CertificateObject += ,('Organization', $CertValueO, '', '')
												$CertificateObject += ,('Organizational Unit', $CertValueOU, '', '')
												$CertificateObject += ,('Email Address', $CertValueE, '', '')
											}
											else
											{
												$CertificateObject += ,('Common Name', $CertCommonName, 'Common Name', $CertIssuerCN)
												$CertificateObject += ,('ISO Country Code', $CertValueC, 'ISO Country Code', $CertIssuerC)
												$CertificateObject += ,('State or Province', $CertValueST, 'State or Province', $CertIssuerST)
												$CertificateObject += ,('Locality', $CertValueL, 'Locality', $CertIssuerL)
												$CertificateObject += ,('Organization', $CertValueO, 'Organization', $CertIssuerO)
												$CertificateObject += ,('Organizational Unit', $CertValueOU, 'Organizational Unit', $CertIssuerOU)
												$CertificateObject += ,('Email Address', $CertValueE, 'Email Address', $CertIssuerE)
											}
											$CertificateObject += ,('SPAN-L', 'Certificate', 'SPAN-R', '')
											$CertificateObject += ,('Not Valid Before', $CertCreateDate, '', '')
											$CertificateObject += ,('Not Valid After', $CertExpiryDate, '', '')
											$CertificateObject += ,('Serial Number', $cert.IE.CertSerialNumber, '', '')
											$CertificateObject += ,('Signature Algorithm', $cert.IE.CertSignatureAlgo, '', '')
											$CertificateObject += ,('Key Length', $cert.IE.CertKeyLength, '', '')
											$CertificateObject += ,('Enhanced Key Usage', $cert.IE.CertExtendedKeyUsage, '', '')
											$CertificateObject += ,('Key Usage', $cert.IE.CertKeyUsageStr, '', '')
											$CertificateSANs = [regex]::replace($cert.IE.CertExtendedSubjectAltName, ', ' , "`n") # Write each SAN on a new line
											$CertificateObject += ,('Subject Alternative Name', $CertificateSANs, '', '')
											$CertificateObject += ,('Verify Status', $cert.IE.CertVerifyStatus, '', '')
											switch ($cert.IE.CertFileType)
											{
												{($_ -eq '1') -or ($_ -eq '4')}
												{
													$CertificateCollectionMy += , ('SBC Certificates', ('SBC Primary Certificate - {0}' -f $CertCommonName), '', $CertificateObject)
												}
												'5'
												{
													$CertificateCollectionMy += , ('SBC Certificates', ('SBC Supplementary Certificates - {0}' -f $CertCommonName), '', $CertificateObject)
												}
												default
												{
													$CertificateCollectionRoot += , ('SBC Certificates', ('Trusted CA Certificate - {0}' -f $CertCommonName), '', $CertificateObject)
												}
											}
											# CertFileType=1 appears to be the SBC's cert. 3 is a Trusted or intermediate cert. I've only ever seen one instance of a 4 & it was the default Sonus device cert.
										}
									}
								}
							}
							# ---- SystemNet ----------
							if ($systemgroup.name -eq 'SystemNet')
							{
								$SystemNodeLevelData = @()
								if ($systemgroup.IE.classname -eq 'SYSTEM_SERV_NET_CFG_IE')
								{
									if ($systemgroup.IE.sysDescription -eq '')
									{
										Write-AtBookmark -bookmark 'sysDescription' -data '<Not Provided>'
									}
									else
									{
										Write-AtBookmark -bookmark 'sysDescription' -data $systemgroup.IE.sysDescription.ToString()
									}

									if ($systemgroup.IE.sysLocation -eq '')
									{
										Write-AtBookmark -bookmark 'sysLocation' -data '<Not Provided>'
									}
									else
									{
										Write-AtBookmark -bookmark 'sysLocation' -data $systemgroup.IE.sysLocation.ToString()
									}
									if ($systemgroup.IE.sysContact -eq '')
									{
										Write-AtBookmark -bookmark 'sysContact' -data '<Not Provided>'
									}
									else
									{
										Write-AtBookmark -bookmark 'sysContact' -data $systemgroup.IE.sysContact.ToString()
									}
									$PrimaryDNSServer = $systemgroup.IE.PrimaryDNSServer #Used in v5+ for the Protocols / IP /DNS / DNS Table
									$SecondaryDNSServer = $systemgroup.IE.SecondaryDNSServer #Used in v5+ for the Protocols / IP /DNS / DNS Table
									$SystemNetValues = @()
									$SystemNetValues += ,('SPAN', 'System Overview', '', '')
									if ($SWeLite)
									{
										$SystemNetValues += ,('SWe Lite ID', $NodeId, '', '')
									}
									else
									{
										$SystemNetValues += ,('Node Serial Number', $NodeId, '', '')
										$SystemNetValues += ,('Hardware ID', $HardwareId, '', '')
										$SystemNetValues += ,('Node-Level Network Settings', (Test-ForNull -LookupTable $NodeLevelNetworkSettingsLookup -value $SystemGroup.IE.useDynamicNetSettings), '', '')
									}
									$SystemNetValues += ,('SPAN-L', 'Host Information', 'SPAN-R', 'Domain Name Service')
									#From here I split the page into LH and RH columns to deal with all the possible different combinations.
									$SystemNetL1 = @()
									$SystemNetL2 = @()
									$SystemNetR1 = @()
									$SystemNetR2 = @()
									#LH:
									$SystemNetL1 += 'Host Name'
									$SystemNetL2 += $systemgroup.IE.NodeName
									$SystemNetL1 += 'Domain Name'
									$SystemNetL2 += $systemgroup.IE.DomainName
									$SystemNetL1 += 'SPAN-L'
									$SystemNetL2 += 'System Information'
									$SystemNetL1 += 'System Description'
									$SystemNetL2 += $systemgroup.IE.sysDescription
									$SystemNetL1 += 'System Location'
									$SystemNetL2 += $systemgroup.IE.sysLocation
									$SystemNetL1 += 'System Contact'
									$SystemNetL2 += $systemgroup.IE.sysContact
									$SystemNetL1 += 'SPAN-L'
									$SystemNetL2 += 'Time Management'
									$SystemNetL1 += 'Time Zone'
									$SystemNetL2 += $SystemGroup.IE.TimeZoneGeoName
									$SystemNetL1 += 'SPAN-L'
									$SystemNetL2 += 'Network Time Protocol'
									if (($SystemGroup.IE.useDynamicNetSettings -eq 1) -and (!$SWeLite))
									{
										$SystemNetL1 += 'Primary NTP Server'
										$SystemNetL2 += 'None'
										$SystemNetL1 += 'Secondary NTP Server'
										$SystemNetL2 += 'None'
									}
									else
									{
										if ($SystemGroup.IE.IsNTP1Enabled -eq '1')
										{
											if (!$SWeLite)
											{
												$SystemNetL1 += 'Use NTP'
												$SystemNetL2 += 'Yes'
											}
											$SystemNetL1 += 'NTP Server'
											$SystemNetL2 += $SystemGroup.IE.NTP1ServerName
											if ($SystemGroup.IE.IsNTP1AuthEnabled -eq '1')
											{
												$SystemNetL1 += 'NTP Server Authentication'
												$SystemNetL2 += 'Enabled'
												$SystemNetL1 += 'NTP Server - MD5 Key ID'
												$SystemNetL2 += $SystemGroup.IE.NTP1ServerKeyID
												$SystemNetL1 += 'NTP MD5 Key'
												$SystemNetL2 += $SystemGroup.IE.NTP1ServerKey
											}
											else
											{
												$SystemNetL1 += 'NTP Server Authentication'
												$SystemNetL2 += 'Disabled'
											}
											$SystemNetL1 += 'SPAN-L'
											$SystemNetL2 += 'NTP Server 2'
											if ($SystemGroup.IE.IsNTP2Enabled -eq '1')
											{
												$SystemNetL1 += 'Use NTP Server 2'
												$SystemNetL2 += 'Yes'
												$SystemNetL1 += 'NTP Server 2'
												$SystemNetL2 += $SystemGroup.IE.NTP2ServerName
												if ($SystemGroup.IE.IsNTP2AuthEnabled -eq '1')
												{
													$SystemNetL1 += 'NTP Server 2 Authentication'
													$SystemNetL2 += 'Enabled'
													$SystemNetL1 += 'NTP Server 2 - MD5 Key ID'
													$SystemNetL2 += $SystemGroup.IE.NTP2ServerKeyID
													$SystemNetL1 += 'NTP Server 2 MD5 Key'
													$SystemNetL2 += $SystemGroup.IE.NTP2ServerKey
												}
												else
												{
													$SystemNetL1 += 'NTP Server Authentication'
													$SystemNetL2 += 'Disabled'
												}
											}
											else
											{
												$SystemNetL1 += 'Use NTP Server 2'
												$SystemNetL2 += 'No'
											}
										}
										else
										{
											$SystemNetL1 += 'Use NTP'
											$SystemNetL2 += 'No'
										}
									}
									if ($SWeLite)
									{
										#System LEDs do not display here
									}
									else
									{
										$SystemNetL1 += 'SPAN-L'
										$SystemNetL2 += 'System LEDs'
										$SystemNetL1 += 'Power LED'
										$SystemNetL2 += '<n/a>'
										$SystemNetL1 += 'Alarm LED'
										$SystemNetL2 += '<n/a>'
										$SystemNetL1 += 'Ready LED'
										$SystemNetL2 += '<n/a>'
										$SystemNetL1 += 'Locator LED'
										$SystemNetL2 += '<n/a>'
									}
									$SystemNetL1 += 'SPAN-L'
									$SystemNetL2 += 'Country Level Information'
									$SystemNetL1 += 'Country Code'
									$SystemNetL2 += Test-ForNull -LookupTable $CountryCodeLookup -value $systemgroup.IE.CountryCode
									#RH:
									if (($SystemGroup.IE.useDynamicNetSettings -eq 1) -and (!$SWeLite))
									{
										$SystemNetR1 += 'Enable DNS Service'
										$SystemNetR2 += Test-ForNull -LookupTable $YesNoLookup -value $systemgroup.IE.IsDNSForwardingEnabled
									}
									else
									{
										if ($systemgroup.IE.PrimaryDNSServer -eq '0.0.0.0')
										{
											$SystemNetR1 += 'Use Primary DNS'
											$SystemNetR2 += 'No'
											if ($SWeLite)
											{
												#Add some padding so the next level of headings align
												$SystemNetR1 += ''
												$SystemNetR2 += ''
											}
											else
											{
												$SystemNetR1 += 'Enable DNS Service'
												$SystemNetR2 += Test-ForNull -LookupTable $YesNoLookup -value $systemgroup.IE.IsDNSForwardingEnabled
											}
											#Add some padding so the next level of headings align
											$SystemNetR1 += ''
											$SystemNetR2 += ''
											$SystemNetR1 += ''
											$SystemNetR2 += ''
											$SystemNetR1 += ''
											$SystemNetR2 += ''
											$SystemNetR1 += ''
											$SystemNetR2 += ''
										}
										else
										{
											$SystemNetR1 += 'Use Primary DNS'
											$SystemNetR2 += 'Yes'
											$SystemNetR1 += 'Primary Server IP'
											$SystemNetR2 += $systemgroup.IE.PrimaryDNSServer
											$SystemNetR1 += 'Primary Source'
											$SystemNetR2 += Test-ForNull -LookupTable $PortToIPAddressLookup -value $systemgroup.IE.PrimaryDNSServerSource
											if ($systemgroup.IE.SecondaryDNSServer -eq '0.0.0.0')
											{
												$SystemNetR1 += 'Use Secondary DNS'
												$SystemNetR2 += 'No'
												if ($SWeLite)
												{
													#Add some padding so the next level of headings align
													$SystemNetR1 += ''
													$SystemNetR2 += ''
												}
												else
												{
													$SystemNetR1 += 'Enable DNS Service'
													$SystemNetR2 += Test-ForNull -LookupTable $YesNoLookup -value $systemgroup.IE.IsDNSForwardingEnabled
												}
												#Add some padding so the next level of headings align
												$SystemNetR1 += ''
												$SystemNetR2 += ''
												$SystemNetR1 += ''
												$SystemNetR2 += ''
											}
											else
											{
												$SystemNetR1 += 'Use Secondary DNS'
												$SystemNetR2 += 'Yes'
												$SystemNetR1 += 'Secondary Server IP'
												$SystemNetR2 += $systemgroup.IE.SecondaryDNSServer
												$SystemNetR1 += 'Secondary Source'
												$SystemNetR2 += Test-ForNull -LookupTable $PortToIPAddressLookup -value $systemgroup.IE.SecondaryDNSServerSource
												if ($SWeLite)
												{
													#Add some padding so the next level of headings align
													$SystemNetR1 += ''
													$SystemNetR2 += ''
												}
												else
												{
													$SystemNetR1 += 'Enable DNS Service'
													$SystemNetR2 += Test-ForNull -LookupTable $YesNoLookup -value $systemgroup.IE.IsDNSForwardingEnabled
												}
												#Add some padding so the next level of headings align
												$SystemNetR1 += ''
												$SystemNetR2 += ''
											}
										}
									}
									if ($SWeLite)
									{
										#DHCP does not display here
									}
									else
									{
										$SystemNetR1 += 'SPAN-R'
										$SystemNetR2 += 'DHCP Server'
										$SystemNetR1 += 'Enable DHCP Server'
										$SystemNetR2 += Test-ForNull -LookupTable $YesNoLookup -value $systemgroup.IE.IsDHCPServerEnabled #This wasn't present prior to 2.1 (?)
									}
									#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array
									if ($SystemNetL1.Count -ge $SystemNetR1.Count)
									{
										$arrayCount = $SystemNetL1.Count
									}
									else
									{
										$arrayCount = $SystemNetR1.Count
									}
									for ($i = 0; $i -lt $arrayCount; $i++)
									{
										if (($SystemNetL1[$i] -eq ''   ) -and ($SystemNetL2[$i] -eq ''   ) -and ($SystemNetR1[$i] -eq ''   ) -and ($SystemNetR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
										if (($SystemNetL1[$i] -eq $null) -and ($SystemNetL2[$i] -eq $null) -and ($SystemNetR1[$i] -eq $null) -and ($SystemNetR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
										$SystemNetValues += ,($SystemNetL1[$i], $SystemNetL2[$i], $SystemNetR1[$i], $SystemNetR2[$i])
									}
									#$SystemNodeLevelData = ,('Node-Level Settings', '', '', $SystemNetValues) - now at the bottom for correct ordering.
								}
							}
							# ---- QoE (new in 3.1) ----------
							if ($systemgroup.name -eq 'QOE')
							{
								$SystemQoEData = @()
								if ($systemgroup.IE.classname -eq 'QOE_IE')
								{
									$QoEColumnTitles = @('QoE Agent Type', 'QoE Server Address', 'QoE Server Port', 'Protocol', 'TLS Profile')
									$QoECollection = @()
									If ($TlsProfileIdLookup.Get_Item($systemgroup.IE.HQGTLSProfileID) -eq '4')
									{
										$QoETlsProfile = $TlsProfileIdLookup.Get_Item($systemgroup.IE.HQGTLSProfileID)
									}
									else
									{
										$QoETlsProfile = '<n/a>' # (The stored value is set to a valid TLS profile even if the protocol is not TLS)
									}
									if ($systemgroup.IE.BranchOfficeGateway -eq '0')
									{
										$QoEObject = @( 'Server', '<n/a>', '<n/a>', '<n/a>', '<n/a>')
									}
									else
									{
										$QoEObject = @( 'Client', $systemgroup.IE.HQGAddress, $systemgroup.IE.HQGPort, $ProtocolLookup.Get_Item($systemgroup.IE.HQGProtocol), $QoETlsProfile)
									}
									$QoECollection += , $QoEObject
									#$SystemQoEData += , ('QoE', '', $QoEColumnTitles, $QoECollection) - now at the bottom for correct ordering.
								}
							}
							# ---- TimingConfig ----------
							if ($systemgroup.name -eq 'TimingConfig')
							{
								$SystemTimingData = @()
								if ($systemgroup.IE.classname -eq 'SYSTEM_TIMING_CONFIG_IE')
								{
									$TimingConfigColumnTitles = @('Clock Source', 'Primary Clock Port', 'Secondary Clock Port')
									$TimingConfigCollection = @()
									if ($systemgroup.IE.TxClockSource -eq '0')
									{
										$PrimaryClockRecoveryPort   = '<n/a>'
										$SecondaryClockRecoveryPort = '<n/a>'
									}
									else
									{
										if (($systemgroup.IE.PrimaryClockRecoveryPort -eq '0.0') -or ($systemgroup.IE.PrimaryClockRecoveryPort -eq ''))
										{
											$PrimaryClockRecoveryPort   = '<n/a>'
										}
										else
										{
											$PrimaryClockRecoveryPort = $systemgroup.IE.PrimaryClockRecoveryPort
										}
										if (($systemgroup.IE.SecondaryClockRecoveryPort -eq '0.0') -or ($systemgroup.IE.SecondaryClockRecoveryPort -eq ''))
										{
											$SecondaryClockRecoveryPort = '<n/a>'
										}
										else
										{
											$SecondaryClockRecoveryPort = $systemgroup.IE.SecondaryClockRecoveryPort
										}
									}
									$TimingConfigObjects = @($TxClockSourceLookup.Get_Item($systemgroup.IE.TxClockSource), $PrimaryClockRecoveryPort, $SecondaryClockRecoveryPort)
									$TimingConfigCollection += , $TimingConfigObjects
									#$SystemTimingData += , ('System Timing', '', $TimingConfigColumnTitles, $TimingConfigCollection) - now at the bottom for correct ordering.
								}
							}
							# ---- Companding Law Config ----------
							if ($systemgroup.name -eq 'CompandingLawConfig')
							{
								$SystemCompandingData = @()
								if ($systemgroup.IE.classname -eq 'SYSTEM_COMPANDING_LAW_CONFIG_IE')
								{
									$CompandingLawConfigColumnTitles = @('Companding Law', '') # Forces this into a H-layout table
									$CompandingLawConfigCollection = @()
									$CompandingLawConfigObject = @($CompandingLawConfigLookup.Get_Item($systemgroup.IE.CompandingLaw))
									$CompandingLawConfigCollection += , $CompandingLawConfigObject
									#$SystemCompandingData += , ('Companding Law Config', '', $CompandingLawConfigColumnTitles, $CompandingLawConfigCollection) - now at the bottom for correct ordering.
								}
							}

							# ---- Media ----------
							if ($systemgroup.name -eq 'Media')
							{
								if ($systemgroup.IE.classname -eq 'MEDIA_SYSTEM_IE')
								{
									#We need to check for some values not present in older firmware:
									$SystemMediaMOHSource = Test-ForNull -LookupTable $MOHSourceLookup -value $systemgroup.IE.MOHSource
									if ($systemgroup.IE.FXSPortSelection -eq $null)
									{
										$SystemMediaFXSPort = '<n/a this rls>'
									}
									else
									{
										if ($systemgroup.IE.MOHSource -eq '0')
										{
											$SystemMediaFXSPort = '<n/a>'
										}
										else
										{
											$SystemMediaFXSPort = $systemgroup.IE.FXSPortSelection
										}
									}
									$MediaSystemCfgTable += ,('SPAN', 'Media System Configuration', '' , '')
									$MediaSystemCfgTable += ,('SPAN-L', 'Port Range', 'SPAN-R', 'Music on Hold')
									#Now build the table - a little different for the SWe Lite:
									if ($SWeLite)
									{
										$MediaSystemCfgTable += ,('Start Port', ($systemgroup.IE.RTPRTCP_PortStart + ' [16384..32767]'), 'Current Music File', $MOHFilename)
										$MediaSystemCfgTable += ,('Number of Port Pairs', ($systemgroup.IE.RTPRTCP_PortCount + ' [1..5000]'), '', '')
										$MediaSystemCfgTable += ,('SPAN-L', '', '', '')
									}
									else
									{
										$MediaSystemCfgTable += ,('Start Port', ($systemgroup.IE.RTPRTCP_PortStart + ' [1024..32767]'), 'Music on Hold Source', $SystemMediaMOHSource)
										if ($SystemMediaMOHSource -eq 'Live')
										{
											$MediaSystemCfgTable += ,('Number of Port Pairs', ($systemgroup.IE.RTPRTCP_PortCount + ' [1..4800]'), 'FXS Port', $SystemMediaFXSPort)
										}
										else
										{
											$MediaSystemCfgTable += ,('Number of Port Pairs', ($systemgroup.IE.RTPRTCP_PortCount + ' [1..4800]'), 'Current Music File', $MOHFilename)
										}
										$MediaSystemCfgTable += ,('SPAN-L', '', '', '')
										$MediaSystemCfgTable += ,('Echo Canceller Type Option', (Test-ForNull -LookupTable $EchoCancellerLookup -value $systemgroup.IE.EchoCancel_LECOption), '', '')
										$MediaSystemCfgTable += ,('Echo Cancel NLP Option', (Test-ForNull -LookupTable $EchoCancelNLPLookup -value $systemgroup.IE.EchoCancel_NLPOption), '', '')
									}
									$MediaSystemCfgTable += ,('Send STUN Packets', (Test-ForNull -LookupTable $EnabledLookup -value $systemgroup.IE.SendStunPackets), '', '')
									$MediaSystemConfigData += , ('Media System Configuration', '', '', $MediaSystemCfgTable)
								}
							}
							# ---- Relay Config ----------
							if ($systemgroup.name -eq 'RelayConfig')
							{
								$RelayData = @()
								if ($systemgroup.IE.classname -eq 'SYSTEM_RELAY_CONFIG_IE')
								{
									$RelayColumnTitles = @('Enabled', 'Relay Config State', 'Analog Sys Relay')
									$RelayCollection = @()
									$RelayObject = @($EnabledLookup.Get_Item($systemgroup.IE.Enabled), $RelayConfigStateLookup.Get_Item($systemgroup.IE.DigitalRelayState), $systemgroup.IE.AnalogSysRelay)
									$RelayCollection += , $RelayObject
									$RelayData += , ('Relay Config', '', $RelayColumnTitles, $RelayCollection)
								}
							}
							# ---- HOSTS ----------
							if ($systemgroup.name -eq 'HOSTS')
							{
								$IPhostsColumnTitles = @('Host Name', 'IP v4 Address', 'Dynamic Refresh')
								$IPhostsCollection = @()
								$IPhosts = $systemgroup.GetElementsByTagName('ID')
								if ($IPhosts.Count -ne 0)
								{
									ForEach ($IPhost in $IPhosts)
									{
										if ($IPhost.IE.classname -eq 'HOSTS_CONF_IE')
										{
											if (($IPhost.IE.IPAddress -ne $null) -and ($IPhost.IE.IPAddress -ne ''))
											{
												$HostIPAddress = $IPhost.IE.IPAddress
											}
											else
											{
												$HostIPAddress = $IPhost.IE.IPV4Address
											}
											$IPhostObject = @($IPhost.IE.HostName, $HostIPAddress, (Test-ForNull -LookupTable $YesNoLookup -value $IPhost.IE.DynamicRefresh))
											$IPhostsCollection += , $IPhostObject
										}
									}
									$DNSData += ,('DNS', 'Hosts', $IPhostsColumnTitles, $IPhostsCollection)
								}
							}
							# ---- Split DNS ----------
							if ($systemgroup.name -eq 'SplitDnsList')
							{
								$SplitDNSColumnTitles = @('DNS Server IP', 'Domain Name', 'Primary Key')
								$SplitDNSCollection = @()
								$SplitDNSList = $systemgroup.GetElementsByTagName('ID')
								if ($SplitDNSList.Count -ne 0)
								{
									ForEach ($SplitDNSEntry in $SplitDNSList)
									{
										if ($SplitDNSEntry.IE.classname -eq 'SPLITDNS_LIST_IE')
										{
											$SplitDNSItems = $SplitDNSEntry.GetElementsByTagName('ID')
											foreach ($SplitDNSItem in $SplitDNSItems)
											{
												if ($SplitDNSItem.IE.classname -eq $null) { continue } # Empty / deleted entry
												if ($SplitDNSItem.IE.classname -eq 'SPLITDNS_CONF_IE')
												{
													$SplitDNSObject = @($SplitDNSItem.IE.DNSServerIP, $SplitDNSItem.IE.DomainName, $SplitDNSItem.Value)
													$SplitDNSCollection += , $SplitDNSObject
												}
											}
										}
									}
									#Now we need to rearrange the entries in the table into their correct sequence:
									if (($SplitDNSList.IE.Sequence -ne '') -and ($SplitDNSList.IE.Sequence -ne $null))
									{
										$sequence = Decode-Sequence -EncodedSequence $SplitDNSList.IE.Sequence -byteCount 2
										$SplitDNSTemp = @()
										foreach ($reindex in $sequence)
										{
											foreach ($tableRow in $SplitDNSCollection)
											{
												if ($tableRow[2] -eq $reindex) # (2 is the 'Index' / 'Primary Key' value)
												{
													$SplitDNSTemp += ,$TableRow
												}
											}
										}
										$SplitDNSCollection = $SplitDNSTemp
									}
									$DNSData += ,('DNS', 'Split DNS', $SplitDNSColumnTitles, $SplitDNSCollection)
								}
								#---------- DNS Table ---------------------------
								# Here we concoct the 'DNS Table' - it's simply the DNS servers added to the Split DNS table
								$DNSTableColumnTitles = @('DNS Server IP', 'Domain Name', 'Configured By')
								$DNSTableCollection = @()
								if ($PrimaryDNSServer -ne '0.0.0.0')
								{
									$DNSTableObject = @($PrimaryDNSServer, 'N/A', 'Node Level Settings')
									$DNSTableCollection += ,$DNSTableObject
								}
								if ($SecondaryDNSServer -ne '0.0.0.0')
								{
									$DNSTableObject = @($SecondaryDNSServer, 'N/A', 'Node Level Settings')
									$DNSTableCollection += ,$DNSTableObject
								}
								foreach ($SplitDNSEntry in $SplitDNSCollection)
								{
									$DNSTableObject = @($SplitDNSEntry[0], $SplitDNSEntry[1], 'Split DNS')
									$DNSTableCollection += ,$DNSTableObject
								}
								$DNSData += ,('DNS', 'DNS Table', $DNSTableColumnTitles, $DNSTableCollection)
							}
						}
					}

					#---------- SECURITY ---------------------------
					'Security'
					{
						$securitygroups = $node.GetElementsByTagName('Token')
						if ($securitygroups.Count -ne 0)
						{
							ForEach ($securitygroup in $securitygroups)
							{
								# ---- TLS Profile ----------
								if ($securitygroup.name -eq 'TLSProfile')
								{
									$TLSCollection = @() #null the collection for each table
									$Tlsprofiles = $securitygroup.GetElementsByTagName('ID')
									ForEach ($Tlsprofile in $Tlsprofiles)
									{
										if ($TlsProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										$TlsProfileTable = @()
										$TlsProfileDescripion = (Fix-NullDescription -TableDescription $TlsProfile.IE.Description -TableValue $TlsProfile.value -TablePrefix 'TLS Profile ID ')
										$TlsProfileIdLookup.Add($TlsProfile.value, $TlsProfileDescripion)
										#We need to check for some values not present in older firmware:
										$TlsProfileValidateServerFQDN = Test-ForNull -LookupTable $EnabledLookup -value $TlsProfile.IE.ValidateServerFQDN
										#Build the table:
										$TlsProfileTable += ,('SPAN-L', 'TLS Profile', '' , '')
										$TlsProfileTable += ,('Description', $TlsProfileDescripion, '' , '')
										$TlsProfileTable += ,('SPAN-L', 'TLS Parameters', '' , '')
										$TlsProfileTable += ,('SPAN-L', 'Common Attributes', '' , '')
										$TlsProfileTable += ,('TLS Protocol', $TlsProtocolLookup.Get_Item($TlsProfile.IE.TLSVersion), '' , '')
										if ($TlsProfile.IE.MutualAuth -eq '0')
										{
											$TlsProfileTable += ,('Mutual Authentication', 'Disabled', '' , '')
										}
										else
										{
											$TlsProfileTable += ,('Mutual Authentication', 'Enabled', '' , '')
										}
										if ($TlsProfile.IE.AllowWeakCiphers -ne $null)
										{
											$TlsProfileTable += ,('Allow Weak Cipher', $EnabledLookup.Get_Item($TlsProfile.IE.AllowWeakCiphers), '' , '')
										}
										$TlsProfileTable += ,('Handshake Inactivity timeout', ($TlsProfile.IE.HandshakeTimeout + ' secs [1..30]'), '' , '')
										$TlsProfileTable += ,('Certificate', (Test-ForNull -LookupTable $CertificateLookup -value $TlsProfile.IE.ClientCertificate), '' , '')
										$TlsProfileTable += ,('SPAN-L', 'Client Attributes', '' , '')
										if ($TlsProfile.IE.ClientCipherSequence -ne $null)
										{
											#New config:
											$TlsCipherSequence = $null
											$TlsClientCipherList = ($TlsProfile.IE.ClientCipherSequence).Split(',') #Value in the file is formatted as '0,2'
											foreach ($TlsClientCipher in $TlsClientCipherList)
											{
												$TlsCipherSequence += ($TlsClientCipherLookupV4.Get_Item($TlsClientCipher) + "`n")
											}
											$TlsCipherSequence = Strip-TrailingCR -DelimitedString $TlsCipherSequence
											$TlsProfileTable += ,('Client Cipher List', $TlsCipherSequence, '' , '')
										}
										else
										{
											$TlsProfileTable += ,('Client Cipher', $TlsClientCipherLookup.Get_Item($TlsProfile.IE.ClientCipher), '' , '')
										}
										if ($TlsProfile.IE.MutualAuth -eq '0')
										{
											if ($TlsProfile.IE.VerifyPeersCertificate -eq '1')
											{
												$TlsProfileTable += ,('Verify Peer Server Certificate', 'Enabled', '' , '')
												$TlsProfileTable += ,('Validate Server FQDN',  $TlsProfileValidateServerFQDN, '' , '')
											}
											else
											{
												$TlsProfileTable += ,('Verify Peer Server Certificate', 'Disabled', '' , '')
											}
										}
										else
										{
											$TlsProfileTable += ,('Validate Server FQDN',  $TlsProfileValidateServerFQDN, '' , '')
										}
										$TlsProfileTable += ,('SPAN-L', 'Server Attribute', '' , '')
										#$TlsProfileTable += ,('Fallback Compatible Mode', (testForNull $EnabledLookup $TlsProfile.IE.FallbackCompatibleMode), '' , '')
										if ($TlsProfile.IE.MutualAuth -eq '1')
										{
											$TlsProfileTable += ,('Validate Client FQDN', $EnabledLookup.Get_Item($TlsProfile.IE.ValidateClientFQDN), '' , '')
										}
										$TLSCollection += , ('TLS Profiles', $TlsProfileDescripion, '', $TlsProfileTable)
									}
									#The $TLSCollection is added to $SecurityData just prior to printing to ensure the correct ordering in the DOCX
								}
								# ---- Bad Actors ----------
								if ($securitygroup.name -eq 'RibbonProtectBadActors')
								{
									$BadActorCollection = @() #null the collection for each table
									$BadActorColumnTitles = @('Actor Type', 'Actor Data', 'Precedence', 'Primary key')
									$BadActorProfiles = $securitygroup.GetElementsByTagName('ID')
									$BadActorPrecedence = ""
									ForEach ($BadActorProfile in $BadActorProfiles)
									{
										if ($BadActorProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										$BadActorType = Test-ForNull -LookupTable $BadActorTypeLookup -value $BadActorProfile.IE.ActorType
										switch ($BadActorType)
										{
											{($_ -eq 'Calling number') -or ($_ -eq 'Called Number')}
											{
												$BadActorPrecedence = "<n/a>"
											}
											default
											{
												$BadActorPrecedence = $BadActorProfile.IE.Precedence
											}
										}
										$BadActorObject = @($BadActorType, $BadActorProfile.IE.ActorData, $BadActorPrecedence, $BadActorProfile.value)
										$BadActorCollection += ,$BadActorObject
									}
								}
							}
						}
					}

					#---------- SNMP ------------------------
					'SNMP'
					{
						$SNMP = $node.GetElementsByTagName('Token')
						if ($SNMP.Count -ne 0)
						{
							ForEach ($SNMPObjects in $SNMP)
							{
								# ---- SNMP Management Addresses ----------
								if ($SNMPObjects.name -eq 'SNMPMgrTable')
								{
									$SNMPManagers = $SNMPObjects.GetElementsByTagName('ID')
									if ($SNMPManagers.Count -ne 0)
									{
										$SNMPManagersCollection = @()
										$SNMPManagerColumnTitles = @('Entry', 'Community', 'Manager Address', 'Community Type', 'Trap Enabled')
										ForEach ($SNMPManager in $SNMPManagers)
										{
											if ($SNMPManager.IE.classname -eq $null) { continue } # Empty / deleted entry
											if ($SNMPManager.IE.classname -eq 'SNMP_MANAGER_ADDR_IE')
											{
												$SNMPObject = @($SNMPManager.value, $SNMPManager.IE.snmpCommunity, $SNMPManager.IE.snmpMgrAddr, $SNMPCommunityTypeLookup.Get_Item($SNMPManager.IE.sCommType), $TrueFalseLookup.Get_Item($SNMPManager.IE.snmpTrapEnable))
												$SNMPManagersCollection += , $SNMPObject
											}
										}
										# Moved to the end to be correctly sequenced:
										#$SNMPData += , ('SNMP Management Addresses', '', $SNMPManagerColumnTitles, $SNMPManagersCollection)
									}
								}

								# ---- Alarms/Events Config ----------
								if ($SNMPObjects.name -eq 'AlarmsEvents')
								{
									$SNMPEvents = $SNMPObjects.GetElementsByTagName('ID')
									if ($SNMPEvents.Count -ne 0)
									{
										$SNMPEventsCollection = @()
										$SNMPEventsColumnTitles = @('ID', 'Type', 'Condition', 'Severity', 'Category')
										ForEach ($SNMPEvent in $SNMPEvents)
										{
											if ($SNMPEvent.IE.classname -eq $null) { continue } # Empty / deleted entry
											if ($SNMPEvent.IE.classname -eq 'ALM_CONFIG_ENTRY_IE')
											{
												$SNMPObject =  @(($SNMPEvent.IE.EventID + '.' + $SNMPEvent.IE.EventSubId), $SNMPTrapTypeLookup.Get_Item($SNMPEvent.IE.EvtType), $SNMPEvent.IE.Condition, $SNMPSeverityLookup.Get_Item($SNMPEvent.IE.aSeverity), $SNMPCategoryLookup.Get_Item($SNMPEvent.IE.aCategory))
												$SNMPEventsCollection += , $SNMPObject
											}
										}
										# Moved to the end to be correctly sequenced:
										#$SNMPData += , ('Alarms/Events Config', '', $SNMPEventsColumnTitles, $SNMPEventsCollection)
									}
								}
							}
						}
					}

					#---------- LOGGING ------------------------
					'Logging'
					{
						$LoggingData = @()
						#Port Mirroring exists in the XML under 'IPRouting', but is exposed to the user as Logging.

						$LoggingElements = $node.GetElementsByTagName('Token')
						ForEach($LoggingElement in $LoggingElements)
						{
							# ---- Local log configuration (2k only) ----------
							if (($LoggingElement.name -eq 'WebUI') -and ($platform -eq 'SBC 2000'))
							{
								if ($LoggingElement.IE.classname -eq $null) { continue } # Empty / deleted entry
								if ($LoggingElement.IE.classname -eq 'LOGGER_WEBUI_DEST')
								{
									$LocalLogServersCollection += , ('Global Log Level', $LogLevelLookup.Get_Item($LoggingElement.IE.DefaultLevel), '', '')
								}
								$LoggingData += , ('Local Log Configuration', '', '', $LocalLogServersCollection)
							}

							# ---- Syslog Servers ----------
							if ($LoggingElement.name -eq 'SyslogServer')
							{
								$RemoteLogServers = $LoggingElement.GetElementsByTagName('ID')
								if ($RemoteLogServers.Count -ne 0)
								{
									$RemoteLogServersCollection = @()
									$RemoteLogServersColumnTitles = @('Global Log Level', 'Log Destination', 'Port', 'Protocol', 'Log Facility', 'Enabled', 'Primary Key')
									ForEach ($RemoteLogServer in $RemoteLogServers)
									{
										if ($RemoteLogServer.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($RemoteLogServer.IE.classname -eq 'LOGGER_SYSLOG_DEST')
										{
											$RemoteLogServersObject = @($LogLevelLookup.Get_Item($RemoteLogServer.IE.DefaultLevel), $LogServerLookup.Get_Item($RemoteLogServer.value), (Test-ForNull -LookupTable $null -value $RemoteLogServer.IE.ServerPort), (Test-ForNull -LookupTable $LoggingProtocolLookup -value $RemoteLogServer.IE.ServerTransport), $SysLogFacilityLookup.Get_Item($RemoteLogServer.IE.SyslogFacility), $EnabledLookup.Get_Item($RemoteLogServer.IE.Enabled), $RemoteLogServer.value)
											$RemoteLogServersCollection += , $RemoteLogServersObject
										}
									}
									$LoggingData += , ('Remote Log Servers', '', $RemoteLogServersColumnTitles, $RemoteLogServersCollection)
								}
							}

							# ---- Subsystems ----------
							if ($LoggingElement.name -eq 'DebugLevels')
							{
								$LoggingSubsystems = $LoggingElement.GetElementsByTagName('ID')
								if ($LoggingSubsystems.Count -ne 0)
								{
									$LoggingSubsystemCollection = @()
									$LoggingSubsystemColumnTitles = @('Subsystem Name', 'Log Level', 'Log Destination', 'Primary Key')
									ForEach ($LoggingSubsystem in $LoggingSubsystems)
									{
										if ($LoggingSubsystem.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($LoggingSubsystem.IE.classname -eq 'LOGGER_LEVEL_SET')
										{
											$LoggerName = $LoggingSubsystem.IE.LoggerName
											$LoggerName = [regex]::replace($LoggerName, 'com.sonus.sbc.' , '') # Sonus firmware builds
											$LoggerName = [regex]::replace($LoggerName, 'com.net.ux.' , '') # NET firmware builds (There's going to be one still in operation somewhere!)
											$LoggingSubsystemObject = @($LoggingSubsystemLookup.Get_Item($LoggerName), $LogLevelLookup.Get_Item($LoggingSubsystem.IE.LogLevel), $LogServerLookup.Get_Item($LoggingSubsystem.IE.LogDestination), $LoggingSubsystem.Value)
											$LoggingSubsystemCollection += , $LoggingSubsystemObject
										}
									}
									$LoggingData += , ('Subsystems', '', $LoggingSubsystemColumnTitles, $LoggingSubsystemCollection)
								}
							}
						}
					}

					#---------- SIGNALING GROUPS ------------------------
					'SignalingGroups'
					{
						$signalinggroups = $node.GetElementsByTagName('Token')
						if ($signalinggroups.Count -ne 0)
						{
							ForEach ($signalinggroup in $signalinggroups)
							{
								# ---- ISDN Signaling Groups ----------
								if ($signalinggroup.name -eq 'ISDN')
								{
									$IsdnGroups = $signalinggroup.GetElementsByTagName('ID')
									if ($IsdnGroups.Count -ne 0)
									{
										ForEach ($IsdnGroup in $IsdnGroups)
										{
											if ($IsdnGroup.IE.classname -eq $null) { continue } # Empty / deleted entry
											$IsdnSGTable = @() #null the collection for each table
											$isgnGroupDescription = Fix-NullDescription -TableDescription $IsdnGroup.IE.Description -TableValue $IsdnGroup.value -TablePrefix 'SG #'
											#Now build the table:
											$IsdnSGTable += ,('SPAN', 'ISDN Signaling Group', '' , '')
											$IsdnSGTable += ,('Description', $isgnGroupDescription, '' , '')
											if ($IsdnGroup.IE.customAdminState -eq $null)
											{
												$IsdnSGTable += ,('Admin State', $EnabledLookup.Get_Item($IsdnGroup.IE.Enabled), '', '')
											}
											else
											{
												$IsdnSGTable += ,('Admin State', $EnabledLookup.Get_Item($IsdnGroup.IE.customAdminState), '', '')
											}
											$IsdnSGTable += ,('SPAN-L', 'Channels and Routing', 'SPAN-R', 'Port and Protocol')
											# Columns Diverge here. Prepare the table, building the L & R columns as independent arrays, then consolidating them together ready for Word.
											$ISDNSgL1 = @()
											$ISDNSgL2 = @()
											$ISDNSgR1 = @()
											$ISDNSgR2 = @()
											#LH Column first:
											switch ($IsdnGroup.IE.IsdnSgDirection)
											{
												'0' #Inbound
												{
													$ISDNSgL1 += 'Direction'
													$ISDNSgL2 += 'Inbound'
													if ($IsdnGroup.IE.IsdnSgRingbackPlayPolicy -ne '2')
													{
														$ISDNSgL1 += 'Tone Table'
														$ISDNSgL2 += $ToneTableLookup.Get_Item($IsdnGroup.IE.ToneTableNumber)
													}
													$ISDNSgL1 += 'Action Set'
													$ISDNSgL2 += $ActionSetLookup.Get_Item($IsdnGroup.IE.ActionsetTableNumber)
													$ISDNSgL1 += 'Call Routing Table'
													$ISDNSgL2 += $CallRoutingTableLookup.Get_Item($IsdnGroup.IE.RouteTableNumber)
												}
												'1' #Outbound
												{
													$ISDNSgL1 += 'Channel Hunting'
													$ISDNSgL2 += $SgHuntMethodLookup.Get_Item($IsdnGroup.IE.IsdnSgHuntMethod)
													$ISDNSgL1 += 'Direction'
													$ISDNSgL2 += 'Outbound'
												}
												'2' #Bidirectional
												{
													$ISDNSgL1 += 'Channel Hunting'
													$ISDNSgL2 += $SgHuntMethodLookup.Get_Item($IsdnGroup.IE.IsdnSgHuntMethod)
													$ISDNSgL1 += 'Direction'
													$ISDNSgL2 += 'Bidirectional'
													if ($IsdnGroup.IE.IsdnSgRingbackPlayPolicy -ne '2')
													{
														$ISDNSgL1 += 'Tone Table'
														$ISDNSgL2 += $ToneTableLookup.Get_Item($IsdnGroup.IE.ToneTableNumber)
													}
													$ISDNSgL1 += 'Action Set'
													$ISDNSgL2 += $ActionSetLookup.Get_Item($IsdnGroup.IE.ActionsetTableNumber)
													$ISDNSgL1 += 'Call Routing Table'
													$ISDNSgL2 += $CallRoutingTableLookup.Get_Item($IsdnGroup.IE.RouteTableNumber)
												}
											}
											$ISDNSgL1 += 'No Channel Available Override'
											$ISDNSgL2 += Test-ForNull -LookupTable $Q850DescriptionLookup -value $IsdnGroup.IE.NoChannelAvailableId
											$ISDNSgL1 += 'Play Inband Message Post-Disconnect'
											if ($IsdnGroup.IE.DelayChannelClearing -eq '0')
											{
												$ISDNSgL2 += 'No'
											}
											else
											{
												if ($IsdnGroup.IE.DelayChannelClearing -eq '1')
												{
													$ISDNSgL2 += 'Yes – Until Timer Expiry'
													$ISDNSgL1 += 'Convert DISC PI=1,8 to Progress'
													$ISDNSgL2 += Test-ForNull -LookupTable $YesNoLookup -value $IsdnGroup.IE.ConvertDisconnectToProgress
													$ISDNSgL1 += 'Maximum Message Play Time'
													$ISDNSgL2 += $IsdnGroup.IE.ChannelClearingDelay + ' [1..300] secs'
												}
												else
												{
													$ISDNSgL1 += 'Convert DISC PI=1,8 to Progress'
													$ISDNSgL2 += Test-ForNull -LookupTable $YesNoLookup -value $IsdnGroup.IE.ConvertDisconnectToProgress
												}
												$ISDNSgL1 += 'Immediate Disconnect on Cause Code'
												if ($IsdnGroup.IE.ImmediateDisconnectOnCausecode -eq $null)
												{
													$ISDNSgL2 += '<n/a this rls>'
												}
												elseif ($IsdnGroup.IE.ImmediateDisconnectOnCausecode -eq '1')
												{
													$ISDNSgL2 += 'Yes'
													$ISDNSgL1 += 'Q.850 Cause Codes'
													$ISDNCauseCodeArray = ($IsdnGroup.IE.CauseCodes).Split(',') #Value in the file is formatted as '16,17'
													$CauseCodeList = ''
													foreach ($ISDNCauseCode in $ISDNCauseCodeArray)
													{
														$CauseCodeList += ($Q850DescriptionLookup.Get_Item($ISDNCauseCode) + "`n")
													}
													$CauseCodeList = Strip-TrailingCR -DelimitedString $CauseCodeList
													$ISDNSgL2 += $CauseCodeList
												}
												else
												{
													$ISDNSgL2 += 'No'
												}
											}
											$ISDNSgL1 += 'Call Setup Response Timer'
											if ($IsdnGroup.IE.TimerSanitySetup -eq $null)
											{
												$ISDNSgL2 += '<n/a this rls>'
											}
											else
											{
												$ISDNSgL2 += ($IsdnGroup.IE.TimerSanitySetup / 1000)
											}
											#Now the RH
											$ISDNSgR1 += 'Port Name'
											$ISDNSgR2 += $IsdnGroup.IE.ApplyToPortName
											$ISDNSgR1 += 'Fractional'
											if (($IsdnGroup.IE.PortChannelCountFull -eq $null) -or ($IsdnGroup.IE.PortChannelCountFull -eq '1'))
											{
												$ISDNSgR2 += 'No'
											}
											else
											{
												$ISDNSgR2 += 'Yes'
												$ISDNSgR1 += 'Channel List Selection'
												$ISDNSgR2 += Consolidate-ChannelList -Channels $IsdnGroup.IE.ApplyToChannelList
											}
											$ISDNSgR1 += 'Switch Variant'
											$ISDNSgR2 += $IsdnSwitchVariantLookup.Get_Item($IsdnGroup.IE.ISDNSGSwitchVariant)
											$ISDNSgR1 += 'ISDN Side'
											$ISDNSgR2 += $IsdnSideOrientationLookup.Get_Item($IsdnGroup.IE.ISDNSGSideOrientation)
											$ISDNSgR1 += 'Play Ringback'
											$ISDNSgR2 += $IsdnRingbackLookup.Get_Item($IsdnGroup.IE.IsdnSgRingbackPlayPolicy)
											switch ($IsdnGroup.IE.IsdnSgSwitchVariant)
											{
												{($_ -eq '4') -or ($_ -eq '5')}
												{
													$ISDNSgR1 += 'Overlap Receive Mode'
													$ISDNSgR2 += $EnabledLookup.Get_Item($IsdnGroup.IE.IsdnSgOverlapRxMode)
													$ISDNSgR1 += 'Overlap Send Mode'
													$ISDNSgR2 += Test-ForNull -LookupTable $EnabledLookup -value $IsdnGroup.IE.IsdnSgOverlapTxMode
												}
												default
												{
													$ISDNSgR1 += 'Service Message Capability'
													$ISDNSgR2 += $EnabledLookup.Get_Item($IsdnGroup.IE.IsdnServiceMessageCapability)
												}
											}
											$ISDNSgR1 += 'Stop Far-End T310'
											$ISDNSgR2 += $EnabledLookup.Get_Item($IsdnGroup.IE.IsdnStopFarEndT310UponTryingFromSip)
											$ISDNSgR1 += 'Indicated Channel'
											$ISDNSgR2 += Test-ForNull -LookupTable $IsdnIndicatedChannelLookup -value $IsdnGroup.IE.IsdnSgIndicatedChannel
											$ISDNSgR1 += 'SPAN-R'
											$ISDNSgR2 += 'Switch-specific Parameters'
											if (($IsdnGroup.IE.IsdnSgSwitchVariant -eq '3') -or ($IsdnGroup.IE.IsdnSgSwitchVariant -eq '4'))
											{
												if ($IsdnGroup.IE.ISDNSGSideOrientation -eq  '0')
												{
													#Only displays if the side is User.
													$ISDNSgR1 += 'Send Calling Name'
													$ISDNSgR2 += $EnabledLookup.Get_Item($IsdnGroup.IE.IsdnAllowCallingNameDisplayToSwitch)
												}
											}
											$ISDNSgR1 += 'Add PI to Setup'
											$ISDNSgR2 += Test-ForNull -LookupTable $IsdnAddPiToSetupLookup -value $IsdnGroup.IE.AddSetupProgress
											$ISDNSgR1 += 'Early Media for PI: 2(Dest not ISDN)'
											$ISDNSgR2 += Test-ForNull -LookupTable $EnabledLookup -value $IsdnGroup.IE.EarlyMediaOnDestNotIsdn
											switch ($IsdnGroup.IE.IsdnSgSwitchVariant)
											{
												'4'
												{
													$ISDNSgR1 += 'Send Facility Message Passthrough'
													$ISDNSgR2 += Test-ForNull -LookupTable $EnabledLookup -value $IsdnGroup.IE.SendFacilityPassthrough
													$ISDNSgR1 += 'Send Redirecting Number in Facility'
													$ISDNSgR2 += Test-ForNull -LookupTable $EnabledLookup -value $IsdnGroup.IE.SendRNInFacility
												}
												'5'
												{
													$ISDNSgR1 += 'Send Facility Message Passthrough'
													$ISDNSgR2 += Test-ForNull -LookupTable $EnabledLookup -value $IsdnGroup.IE.SendFacilityPassthrough
												}
												{($_ -eq '0') -or ($_ -eq '1')}
												{
													$ISDNSgR1 += 'Add Network Specific Facilities to Setup'
													$ISDNSgR2 += $EnabledLookup.Get_Item($IsdnGroup.IE.AddNetworkSpecificFacility)
													if ($IsdnGroup.IE.AddNetworkSpecificFacility -eq '1')
													{
														$ISDNSgR1 += 'NSF Information Value'
														$ISDNSgR2 += $IsdnNsfInfoLookup.Get_Item($IsdnGroup.IE.NSFInformation)
														$ISDNSgR1 += 'NSF Identification Value'
														$ISDNSgR2 += $IsdnNsfIdLookup.Get_Item($IsdnGroup.IE.NSFIdentification)
													}
												}
											}
											switch ($IsdnGroup.IE.IsdnSgSwitchVariant)
											{
												'5'
												{
													$ISDNSgR1 += 'ASN.1 Protocol Identifier'
													$ISDNSgR2 += $IsdnAsn1ProtocolLookup.Get_Item($IsdnGroup.IE.ProtocolIdentifier)
													$ISDNSgR1 += 'ASN.1 Numbering Space'
													$ISDNSgR2 += $IsdnAsn1NumberingLookup.Get_Item($IsdnGroup.IE.AsnNumberingSpace)
													$ISDNSgR1 += 'Include NFE and IA PDU'
													$ISDNSgR2 += $EnabledLookup.Get_Item($IsdnGroup.IE.IncludeNFEIAPDU)
												}
												default
												{
													$ISDNSgR1 += 'Include Channel Interface ID'
													$ISDNSgR2 += Test-ForNull -LookupTable $EnabledLookup -value $IsdnGroup.IE.IncludeInterfaceIdentifier
													$ISDNSgR1 += 'Channel Number Bit'
													$ISDNSgR2 += Test-ForNull -LookupTable $IsdnChannelNumberBitLookup -value $IsdnGroup.IE.ChannelNumberBit
												}
											}
											#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array
											if ($ISDNSgL1.Count -ge $ISDNSgR1.Count)
											{
												$arrayCount = $ISDNSgL1.Count
											}
											else
											{
												$arrayCount = $ISDNSgR1.Count
											}
											for ($i = 0; $i -lt $arrayCount; $i++)
											{
												if (($ISDNSgL1[$i] -eq ''   ) -and ($ISDNSgL2[$i] -eq ''   ) -and ($ISDNSgR1[$i] -eq ''   ) -and ($ISDNSgR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
												if (($ISDNSgL1[$i] -eq $null) -and ($ISDNSgL2[$i] -eq $null) -and ($ISDNSgR1[$i] -eq $null) -and ($ISDNSgR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
												$IsdnSGTable += ,($ISDNSgL1[$i], $ISDNSgL2[$i], $ISDNSgR1[$i], $ISDNSgR2[$i])
											}
											$IsdnSGTable += ,('SPAN-L', 'Timeout / Timer Settings', '' , '')
											$IsdnSGTable += ,('T301', $IsdnGroup.IE.IsdnSgTimerT301, '' , '')
											$IsdnSGTable += ,('T302', $IsdnGroup.IE.IsdnSgTimerT302, '' , '')
											$IsdnSGTable += ,('T303', $IsdnGroup.IE.IsdnSgTimerT303, '' , '')
											$IsdnSGTable += ,('T305', $IsdnGroup.IE.IsdnSgTimerT305, '' , '')
											$IsdnSGTable += ,('T308', $IsdnGroup.IE.IsdnSgTimerT308, '' , '')
											$IsdnSGTable += ,('T309', $IsdnGroup.IE.IsdnSgTimerT309, '' , '')
											$IsdnSGTable += ,('T310', $IsdnGroup.IE.IsdnSgTimerT310, '' , '')
											$IsdnSGTable += ,('T313', $IsdnGroup.IE.IsdnSgTimerT313, '' , '')
											$IsdnSGTable += ,('T314', $IsdnGroup.IE.IsdnSgTimerT314, '' , '')
											$IsdnSGTable += ,('T316', $IsdnGroup.IE.IsdnSgTimerT316, '' , '')
											$IsdnSGTable += ,('T322', $IsdnGroup.IE.IsdnSgTimerT322, '' , '')
											$SGData += , ('ISDN', $isgnGroupDescription, '', $IsdnSGTable)
										}
									}
								}

								# ---- SIP Signaling Groups ----------
								if ($signalinggroup.name -eq 'SIP')
								{
									$sipgroups = $signalinggroup.GetElementsByTagName('ID')
									if ($SIPgroups.Count -ne 0)
									{
										ForEach ($SIPgroup in $SIPgroups)
										{
											if ($SIPgroup.IE.classname -eq $null) { continue } # Empty / deleted entry
											if ([int]$SIPgroup.value -lt 50000)
											{
												#It's a "proper" SIP Signaling group (as distinct from a v9.0.0+ SIP *recorder* group, which we'll handle in the "else" below)
												$SipSGTable = @() #null the collection for each table
												#Consolidate all of the host and mask values (old versions had them separated, newer firmware they're all comma-separated together in the one field):
												$SIPGroupFederationIP = ''
												$SIPGroupRemoteHosts = ''
												$SIPGroupRemoteMasks = ''
												if ($SIPgroup.IE.RemoteHosts -eq $null)
												{
													#It's O-L-D firmware. Each one of up to 6 hosts and masks is a separate element
													for ($i = 1 ; $i -le 6 ; $i++)
													{
														if (($SIPGroup.IE.('RemoteHost' + $i) -eq '') -and ($SIPGroup.IE.('RemoteMask' + $i) -eq '')) { continue } #null entry
														if ($SIPGroup.IE.('RemoteMask' + $i) -eq '')
														{
															$SIPGroupFederationIP += $SIPGroup.IE.('RemoteHost' + $i) + " / <n/a>`n"
														}
														else
														{
															$SIPGroupFederationIP += $SIPGroup.IE.('RemoteHost' + $i) + ' / ' + $SIPGroup.IE.('RemoteMask' + $i) + "`n"
														}
													}
												}
												else
												{
													$SIPGroupRemoteHosts = ($SIPgroup.IE.RemoteHosts).Split(',')
													$SIPGroupRemoteMasks = ($SIPgroup.IE.RemoteMasks).Split(',')
													for ($i = 0 ; $i -le ($SIPGroupRemoteHosts.Count-1); $i++)
													{
														if (($SIPGroupRemoteHosts[$i] -eq '') -and ($SIPGroupRemoteMasks[$i] -eq '')) { continue} #Skip a null entry
														if ($SIPGroupRemoteMasks[$i] -eq '')
														{
															$SIPGroupFederationIP += $SIPGroupRemoteHosts[$i] + " / <n/a>`n"
														}
														else
														{
															$SIPGroupFederationIP += $SIPGroupRemoteHosts[$i] + ' / ' + $SIPGroupRemoteMasks[$i] + "`n"
														}
													}
												}
												$SIPGroupFederationIP = Strip-TrailingCR -DelimitedString $SIPGroupFederationIP
												if ($SIPGroupFederationIP -eq '') { $SIPGroupFederationIP = '-- Table is empty --' }
												$SIPgroupDescription = Fix-NullDescription -TableDescription $SIPgroup.IE.Description -TableValue $SIPgroup.value -TablePrefix 'SG #'
												#Now build the table:
												$SipSGTable += ,('SPAN', 'SIP Signaling Group', '' , '')
												$SipSGTable += ,('Description', $SIPgroupDescription, '' , '')
												if ($SIPgroup.IE.customAdminState -eq $null)
												{
													$SipSGTable += ,('Admin State', $EnabledLookup.Get_Item($SIPgroup.IE.Enabled), '', '')
												}
												else
												{
													$SipSGTable += ,('Admin State', $EnabledLookup.Get_Item($SIPgroup.IE.customAdminState), '', '')
												}
												$SipSGTable += ,('SPAN-L', 'SIP Channels and Routing', 'SPAN-R', 'Media Information')
												#From here I split the page into LH and RH columns to deal with all the possible different combinations.
												$SIPSgL1 = @()
												$SIPSgL2 = @()
												$SIPSgR1 = @()
												$SIPSgR2 = @()
												#LH:
												$SIPSgL1 += 'Action Set'
												$SIPSgL2 += $ActionSetLookup.Get_Item($SIPgroup.IE.ActionSetTableID)
												$SIPSgL1 += 'Call Routing Table'
												$SIPSgL2 += $CallRoutingTableLookup.Get_Item($SIPgroup.IE.RouteTableID)
												$SIPSgL1 += 'Channels'
												$SIPSgL2 += $SIPgroup.IE.Channels
												$SIPSgL1 += 'SIP profile'
												$SIPSgL2 += $SipProfileIdLookup.Get_Item($SIPgroup.IE.ProfileID)
												# ----- OK, table handling gets a little ugly here depending upon the value of 'SIP MODE'
												switch ($SIPgroup.IE.Monitor)
												{
													'3' # 'Basic Call'
													{
														$SIPSgL1 += 'SIP Mode'
														$SIPSgL2 += 'Basic Call'
														$SIPSgL1 += 'Agent Type'
														if ($SIPgroup.IE.AgentType -eq $null)
														{
															$SIPSgL2 += '<n/a this rls>'
														}
														else
														{
															switch ($SIPgroup.IE.AgentType)
															{
																'0'
																{
																	$SIPSgL2 += 'Back-to-Back User Agent'
																	if ($SWeLite)
																	{
																		# We don't show Interop Mode for B2B-AU
																	}
																	else
																	{
																		$SIPSgL1 += 'Interop Mode'
																		if ($SIPgroup.IE.InteropMode -eq $null)
																		{
																			$SIPSgL2 += '<n/a this rls>'
																		}
																		else
																		{
																			switch ($SIPgroup.IE.InteropMode)
																			{
																				'0'
																				{
																					$SIPSgL2 += 'Standard'
																				}
																				'2'
																				{
																					$SIPSgL2 += 'Office 365'
																					$SIPSgL1 += 'Office365 User Domain Suffix'
																					$SIPSgL2 += $SIPgroup.IE.Office365FQDN
																				}
																				'3'
																				{
																					$SIPSgL2 += 'Office 365 w/AD PBX'
																					$SIPSgL1 += 'AD Attribute'
																					$SIPSgL2 += $SIPgroup.IE.ADAttribute
																					$SIPSgL1 += 'AD Update Frequency'
																					$SIPSgL2 += $SIPgroup.IE.ADUpdateFrequency + ' [1..30] days'
																					$SIPSgL1 += 'AD First Update Time'
																					$SIPSgL2 += $SIPgroup.IE.ADFirstUpdateTime + ' [hh:mm:ss]'
																					$SIPSgL1 += 'Office365 User Domain Suffix'
																					$SIPSgL2 += $SIPgroup.IE.Office365FQDN
																				}
																				default
																				{
																					$SIPSgL2 += '<Unhandled Value>'
																				}
																			}
																		}
																	}
																}
																'1'
																{
																	$SIPSgL2 += 'Access Mode'
																	$SIPSgL1 += 'Interop Mode'
																	$SIPSgL2 += Test-ForNull -LookupTable $SipSgInteropModeLookup -value $SIPgroup.IE.InteropMode
																	$SIPSgL1 += 'Registrant TTL'
																	$SIPSgL2 += $SIPgroup.IE.RegistrantTTL
																}
																default { $SIPSgL2 += '<Unhandled Value>' }
															}
														}
														$SIPSgL1 += 'SIP Server Table'
														$SIPSgL2 += $SIPServerTablesLookup.Get_Item($SIPgroup.IE.ServerClusterID)
													}
													'1' # 'Fwd Reg after local processing'
													{
														$SIPSgL1 += 'SIP Mode'
														$SIPSgL2 += 'Fwd Reg. After Local Processing'
														$SIPSgL1 += 'Registrar'
														$SIPSgL2 += $SIPRegistrarsLookup.Get_Item($SIPgroup.IE.RegistrarID)
														$SIPSgL1 += 'Registrar Min TTL'
														$SIPSgL2 += $SIPgroup.IE.RegistrarTTL
														$SIPSgL1 += 'Outbound Registrant TTL'
														$SIPSgL2 += $SIPgroup.IE.OutboundRegistrarTTL
														$SIPSgL1 += 'SIP Server Table'
														$SIPSgL2 += $SIPServerTablesLookup.Get_Item($SIPgroup.IE.ServerClusterID)
													}
													'2' # 'Local Registrar'
													{
														$SIPSgL1 += 'SIP Mode'
														$SIPSgL2 += 'Local Registrar'
														$SIPSgL1 += 'Registrar'
														$SIPSgL2 += $SIPRegistrarsLookup.Get_Item($SIPgroup.IE.RegistrarID)
														$SIPSgL1 += 'Agent Type'
														if ($SIPgroup.IE.AgentType -eq $null)
														{
															$SIPSgL2 += '<n/a this rls>'
														}
														else
														{
															switch ($SIPgroup.IE.AgentType)
															{
																'0'
																{
																	$SIPSgL2 += 'Back-to-Back User Agent'
																}
																'1'
																{
																	$SIPSgL2 += 'Access Mode'
																	$SIPSgL1 += 'Interop Mode'
																	$SIPSgL2 += Test-ForNull -LookupTable $SipSgInteropModeLookup -value $SIPgroup.IE.InteropMode
																}
																default
																{
																	$SIPSgL2 += '<Unhandled Value>'
																}
															}
														}
														$SIPSgL1 += 'Registrar Min TTL'
														$SIPSgL2 += Test-ForNull -LookupTable $null -value $SIPgroup.IE.RegistrarTTL
													}
												}
												$SIPSgL1 += 'Load Balancing'
												$SIPSgL2 += $SipLoadBalancingLookup.Get_Item($SIPgroup.IE.ServerSelection)
												$SIPSgL1 += 'Channel Hunting'
												$SIPSgL2 += $SgHuntMethodLookup.Get_Item($SIPgroup.IE.HuntMethod)
												$SIPSgL1 += 'Notify Lync CAC Profile'
												$SIPSgL2 += Test-ForNull -LookupTable $EnableLookup -value $SIPgroup.IE.NotifyCACProfile
												if (($SIPgroup.IE.Monitor -eq '2') -and ($SIPgroup.IE.AgentType -eq '1'))
												{
													#'Challenge request' is hidden if (in Rls 4+) SIP Mode = 'Local Registrar' and Agent Type = 'Access Mode'
												}
												else
												{
													if ($SIPgroup.IE.ChallengeRequest -eq '0')
													{
														$SIPSgL1 += 'Challenge Request'
														$SIPSgL2 += 'Disable'
													}
													else
													{
														$SIPSgL1 += 'Challenge Request'
														$SIPSgL2 += 'Enable'
														$SIPSgL1 += 'Authorization Realm'
														$SIPSgL2 += $SIPgroup.IE.AuthorizationRealm
														$SIPSgL1 += 'Local/Pass-thru Auth Table'
														$SIPSgL2 += $SIPAuthorisationTableLookup.Get_Item($SIPgroup.IE.ProxyAuthorizationTableID)
														$SIPSgL1 += 'Nonce Expiry'
														if ($SIPgroup.IE.NonceLifetime -eq 0)
														{
															$SIPSgL2 += 'Forever'
														}
														else
														{
															$SIPSgL2 += 'Limited'
															$SIPSgL1 += 'Nonce Lifetime'
															$SIPSgL2 += $SIPgroup.IE.NonceLifetime
														}
													}
												}
												$SIPSgL1 += 'Outbound Proxy IP/FQDN'
												$SIPSgL2 += $SIPgroup.IE.OutboundProxy
												if (($SIPgroup.IE.ProxyIpVersion -eq $null) -or ($SIPgroup.IE.OutboundProxy -eq ''))
												{
													# Don't show 'Proxy IP Version'
												}
												else
												{
													try
													{
														if ([ipaddress]$SIPgroup.IE.OutboundProxy) {}
														# It's an IP address: Don't show 'Proxy IP Version'
													}
													catch
													{
														#It's not a valid IP, or it's a hostname
														$SIPSgL1 += 'Proxy IP Version'
														$SIPSgL2 += Test-ForNull -LookupTable $IpVersionLookup -value $SIPgroup.IE.ProxyIpVersion
													}
												}
												$SIPSgL1 += 'Outbound Proxy Port'
												if ($SIPgroup.IE.OutboundProxyPort -eq '0')
												{
													$SIPSgL2 += ''
												}
												else
												{
													$SIPSgL2 += ($SIPgroup.IE.OutboundProxyPort + ' [1024..65535]')
												}
												if ($SWeLite)
												{
													#Doesn't show
												}
												else
												{
													$SIPSgL1 += 'No Channel Available Override'
													$SIPSgL2 += Test-ForNull -LookupTable $Q850DescriptionLookup -value $SIPgroup.IE.NoChannelAvailableId
												}
												$SIPSgL1 += 'Call Setup Response Timer'
												if ($SIPgroup.IE.TimerSanitySetup -eq $null)
												{
													$SIPSgL2 += '<n/a this rls>'
												}
												else
												{
													$SIPSgL2 += (($SIPgroup.IE.TimerSanitySetup / 1000).ToString() + ' [180..750] secs')
												}
												$SIPSgL1 += 'Call Proceeding Timer'
												if ($SIPgroup.IE.TimerCallProceeding -eq $null)
												{
													$SIPSgL2 += '<n/a this rls>'
												}
												else
												{
													$SIPSgL2 += (($SIPgroup.IE.TimerCallProceeding / 1000).ToString() + ' [24..750] secs')
												}
												if ($SWeLite)
												{
													# QoE Reporting does not show
												}
												else
												{
													$SIPSgL1 += 'QoE Reporting'
													$SIPSgL2 += Test-ForNull -LookupTable $EnabledLookup -value $SIPgroup.IE.QoEReporting
												}
												#v5.0: SYM-19541 RegisterKeepAlive in SIP Signaling Group should be displayed only if AgentType=B2BUA
												# We also need to have SIP Mode either 'Basic Call' or 'Forward Reg' (i.e. not 'Local Registrar')
												if (($SIPgroup.IE.AgentType -eq '0') -and ($SIPgroup.IE.Monitor -ne '2'))
												{
													$SIPSgL1 += 'Use Register as Keep Alive'
													$SIPSgL2 += Test-ForNull -LookupTable $EnableLookup -value $SIPgroup.IE.RegisterKeepAlive
												}
												$SIPSgL1 += 'Forked Call Answered Too Soon'
												if ($SIPgroup.IE.RelOnQckConnect -eq $null)
												{
													$SIPSgL2 += '<n/a this rls>'
												}
												else
												{
													if ($SIPgroup.IE.RelOnQckConnect -eq '0')
													{
														$SIPSgL2 += 'Disable'
													}
													else
													{
														$SIPSgL2 += 'Enable'
														$SIPSgL1 += 'Answer Too Soon Timer'
														$SIPSgL2 += (($SIPgroup.IE.RelOnQckConnectTimer).ToString() + ' [1..5000] ms')
													}
												}
												$SIPSgL1 += 'SPAN-L'
												$SIPSgL2 += 'SIP Recording'
												$SIPSgL1 += 'SIP Recording Status'
												$SIPSgL2 += Test-ForNull -LookupTable $EnabledLookup -value $SIPgroup.IE.SipRecordingStatus
												if ($SIPgroup.IE.SipRecordingStatus -eq 1)
												{
													$SIPSgL1 += "SIP Recorder"
													$SIPSgL2 += $SgTableLookup.Get_Item($SIPgroup.IE.SipRecorder)
												}
												#RH:
												if ($SIPgroup.IE.RTPDirectMode -ne $null)
												{
													# From v5.0 (and with the arrival of RTPDirectMode) the display changes:
													$AudioFaxStreamMode = ''
													if ($SWeLite)
													{
														$SIPSgR1 += 'Supported Audio Modes'
													}
													else
													{
														$SIPSgR1 += 'Supported Audio/Fax Modes'
													}
													if ($SIPgroup.IE.RTPMode -eq '1') { $AudioFaxStreamMode += "DSP`n" }
													if ($SIPgroup.IE.RTPProxyMode  -eq '1') { $AudioFaxStreamMode += "Proxy`n" }
													if ($SIPgroup.IE.RTPDirectMode -eq '1') { $AudioFaxStreamMode += "Direct`n" }
													if ($SWeLite)
													{
														if ($SIPgroup.IE.RTPProxySrtpMode -eq '1') { $AudioFaxStreamMode += "Proxy with Local SRTP`n" }
													}
													$SIPSgR2 += Strip-TrailingCR -DelimitedString $AudioFaxStreamMode
												}
												else
												{
													# Pre-v5.0 layout:
													$SIPSgR1 += 'Audio/Fax Stream Proxy Mode'
													$SIPSgR2 += Test-ForNull -LookupTable $EnabledLookup -value $SIPgroup.IE.RTPProxyMode
													$SIPSgR1 += 'Audio/Fax Stream DSP Mode'
													$SIPSgR2 += Test-ForNull -LookupTable $EnabledLookup -value $SIPgroup.IE.RTPMode
												}
												$SIPSgR1 += 'Supported Video/Application Modes'
												if ($SWeLite)
												{
													$SIPSgR2 += 'Disabled'
												}
												else
												{
													if ($LicencedForVideo -eq $null)
													{
														# This will be null if we're reading from an XML as there's no licencing info available
														$SIPSgR2 += '<Unknown>'
													}
													else
													{
														if ($LicencedForVideo)
														{
															$VideoStreamMode = ''
															if ($SIPgroup.IE.VideoProxyMode  -eq '1') { $VideoStreamMode += "Proxy`n"  }
															if ($SIPgroup.IE.VideoDirectMode -eq '1') { $VideoStreamMode += "Direct`n" }
															if ($VideoStreamMode -eq '') { $VideoStreamMode = "[None]" }
															$SIPSgR2 += Strip-TrailingCR -DelimitedString $VideoStreamMode
														}
														else
														{
															$SIPSgR2 += 'Disabled'
														}
													}
												}
												if (($SIPgroup.IE.RTPMode -eq $null) -or ($SIPgroup.IE.RTPMode -eq '1'))
												{
													#If RTP DSP Mode is Disabled, hide all of these values (Part 1):
													$SIPSgR1 += 'Media List ID'
													$SIPSgR2 += $MediaListProfileLookup.Get_Item($SIPgroup.IE.MediaConfigID)
												}
												if ($SWeLite -and $SIPgroup.IE.RTPProxySrtpMode -eq '1')
												{
													$SIPSgR1 += 'Proxy Local SRTP Crypto Profile ID'
													$SIPSgR2 += $SDESMediaCryptoProfileLookup.Get_Item($SIPgroup.IE.CryptoProfileID)
												}
												if (($SIPgroup.IE.RTPMode -eq $null) -or ($SIPgroup.IE.RTPMode -eq '1'))
												{
													$SIPSgR1 += 'Play Ringback'
													$SIPSgR2 += $SipRingbackLookup.Get_Item($SIPgroup.IE.RingBack)
													if ($SIPgroup.IE.RingBack -ne 2)
													{
														$SIPSgR1 += 'Tone Table'
														$SIPSgR2 += $ToneTableLookup.Get_Item($SIPgroup.IE.ToneTableID)
													}
													$SIPSgR1 += 'Play Congestion Tone'
													$SIPSgR2 += Test-ForNull -LookupTable $EnableLookup -value $SIPgroup.IE.PlayCongestionTone
													$SIPSgR1 += 'Early 183'
													$SIPSgR2 += $EnableLookup.Get_Item($SIPgroup.IE.Early183)
												}
												elseif ($SWeLite)
												{
													#So as at v7.0, if you DON'T have 'DSP' as a media type, the Lite still gives you a Tone Table option:
													$SIPSgR1 += 'Tone Table'
													$SIPSgR2 += $ToneTableLookup.Get_Item($SIPgroup.IE.ToneTableID)
												}
												# But everyone gets 'Allow Refresh DSP':
												$SIPSgR1 += 'Allow Refresh SDP'
												$SIPSgR2 += Test-ForNull -LookupTable $EnableLookup -value $SIPgroup.IE.AllowRefreshSDP
												if (($SIPgroup.IE.RTPMode -eq $null) -or ($SIPgroup.IE.RTPMode -eq '1'))
												{
													#If RTP DSP Mode is Disabled, hide all of these values (Part 2):
													$SIPSgR1 += 'Music On Hold'
													$SIPSgR2 += Test-ForNull -LookupTable $SipSgMOHLookup -value $SIPgroup.IE.SGLevelMOHService
												}
												$SIPSgR1 += 'RTCP Multiplexing'
												$SIPSgR2 += Test-ForNull -LookupTable $EnableLookup -value $SIPgroup.IE.RTCPMultiplexing
												$SIPSgR1 += 'SPAN-R'
												$SIPSgR2 += 'Mapping Tables'
												$SIPSgR1 += 'SIP to Q.850 Override Table'
												$SIPSgR2 += $SIPToQ850TableLookup.Get_Item($SIPgroup.IE.SIPtoQ850_TableID)
												$SIPSgR1 += 'Q.850 to SIP Override Table'
												$SIPSgR2 += $Q850ToSIPTableLookup.Get_Item($SIPgroup.IE.Q850toSIP_TableID)
												$SIPSgR1 += 'Pass-thru Peer SIP Response Code'
												$SIPSgR2 += Test-ForNull -LookupTable $EnableLookup -value $SIPgroup.IE.PassthruPeerSIPRespCode
												if ($SIPgroup.IE.ServerSelection -eq '4')
												{
													$SIPSgR1 += 'SIP Failover Cause Codes'
													$SipSgResponseCodes = $null
													$SipResponseCodesList = ($SIPgroup.IE.SipResponseCodes).Split(',') #Value in the file is formatted as '0,2'
													foreach ($SipResponseCode in $SipResponseCodesList)
													{
														$SipSgResponseCodes += ($SIPDescriptionLookup.Get_Item($SipResponseCode) + "`n")
													}
													$SIPSgR2 += Strip-TrailingCR -DelimitedString $SipSgResponseCodes
												}
												$SIPSgR1 += 'SPAN-R'
												$SIPSgR2 += 'SIP IP Details'
												$SIPSgR1 += 'Teams Local Media Optimization'
												$SIPSgR2 += Test-ForNull -LookupTable $EnableLookup -value $SIPgroup.IE.MediaOptimization
												if ($SIPgroup.IE.NATTraversalType -eq '0') #Outbound NAT
												{
													$SIPSgR1 += 'Signaling/Media Source IP'
													$SIPSgR2 += $PortToIPAddressLookup.Get_Item($SIPgroup.IE.NetInterfaceSignaling)
												}
												else
												{
													$SIPSgR1 += 'Signaling/Media Private IP'
													$SIPSgR2 += $PortToIPAddressLookup.Get_Item($SIPgroup.IE.NetInterfaceSignaling)
												}
												if ($SIPgroup.IE.MediaOptimization -eq '1')
												{
													$SIPSgR1 += 'Private Media Source IP'
													$SIPSgR2 += $PortToIPAddressLookup.Get_Item($SIPgroup.IE.PrivateMediaSourceIp)
												}
												$SIPSgR1 += 'Signaling DSCP'
												$SIPSgR2 += Test-ForNull -LookupTable $null -value $SIPgroup.IE.DSCP
												if (($SIPgroup.IE.Monitor -eq '3') -and ($SIPgroup.IE.AgentType -eq '0'))
												{
													# "If SIP Mode = Basic Call & Agent Type = B2BUA"
													$SIPSgR1 += 'SPAN-R'
													$SIPSgR2 += 'NAT Traversal'
													$SIPSgR1 += 'ICE Support'
													$SIPSgR2 += Test-ForNull -LookupTable $EnabledLookup -value $SIPgroup.IE.ICESupport
													if ($SIPgroup.IE.ICESupport -eq '1')
													{
														$SIPSgR1 += 'ICE Mode'
														$SIPSgR2 += Test-ForNull -LookupTable $ICEModeLookup -value $SIPgroup.IE.ICEMode
													}
												}
												$SIPSgR1 += 'SPAN-R'
												$SIPSgR2 += 'Static NAT - Outbound'
												if ($SIPgroup.IE.NATTraversalType -eq '0')
												{
													$SIPSgR1 += 'Outbound NAT Traversal'
													$SIPSgR2 += 'None'
												}
												else
												{
													$SIPSgR1 += 'Outbound NAT Traversal'
													$SIPSgR2 += 'Static NAT'
													$SIPSgR1 += 'NAT Public IP (Signaling/Media)'
													$SIPSgR2 += $SIPgroup.IE.NATPublicIPAddress
												}
												$SIPSgR1 += 'SPAN-R'
												$SIPSgR2 += 'Static NAT - Inbound'
												$SIPSgR1 += 'Detection'
												if ($SIPgroup.IE.InboundNATTraversalDetection -eq $null)
												{
													$SIPSgR2 += '<n/a this rls>'
												}
												else
												{
													if ($SIPgroup.IE.InboundNATTraversalDetection -eq '0')
													{
														$SIPSgR2 += 'Disabled'
													}
													else
													{
														$SIPSgR2 += 'Enabled'
														$SIPSgR1 += 'Qualified Prefixes Table'
														$SIPSgR2 += $SIPNATPrefixesLookup.Get_Item($SIPgroup.IE.InboundNATQualifiedPrefixesTableID)
														$SIPSgR1 += 'Secure Media Latching'
														if ($SIPgroup.IE.InboundSecureNATMediaLatching -eq '0')
														{
															$SIPSgR2 += 'Disabled'
														}
														else
														{
															$SIPSgR2 += 'Enabled'
															$SIPSgR1 += 'Secure Media Netmask'
															$SIPSgR2 += $SIPgroup.IE.InboundSecureNATMediaPrefix
														}
														$SIPSgR1 += 'Registrar Max. TTL Enabled'
														if ($SIPgroup.IE.InboundNATPeerRegistrarMaxEnabled -eq '0')
														{
															$SIPSgR2 += 'No'
														}
														else
														{
															$SIPSgR2 += 'Yes'
															$SIPSgR1 += 'Registrar Max. TTL'
															$SIPSgR2 += ($SIPgroup.IE.InboundNATPeerRegistrarMaxTTL + ' [30..86400] secs')
														}
													}
												}
												#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array
												if ($SIPSgL1.Count -ge $SIPSgR1.Count)
												{
													$arrayCount = $SIPSgL1.Count
												}
												else
												{
													$arrayCount = $SIPSgR1.Count
												}
												for ($i = 0; $i -lt $arrayCount; $i++)
												{
													if (($SIPSgL1[$i] -eq ''   ) -and ($SIPSgL2[$i] -eq ''   ) -and ($SIPSgR1[$i] -eq ''   ) -and ($SIPSgR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
													if (($SIPSgL1[$i] -eq $null) -and ($SIPSgL2[$i] -eq $null) -and ($SIPSgR1[$i] -eq $null) -and ($SIPSgR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
													$SipSGTable += ,($SIPSgL1[$i], $SIPSgL2[$i], $SIPSgR1[$i], $SIPSgR2[$i])
												}
												$SipSGTable += ,('SPAN-L','Listen Ports', 'SPAN-R', 'Federated IP/FQDN')
												#Build the rows here - Listen Ports first
												$SIPListenList = ''
												for ($i = 1; $i -le 6; $i++)
												{
													if ($SIPgroup.IE.('ListenPort_' + $i) -ne '0') #We have a valid entry.
													{
															$SIPListenList += ("{0} : {1} : {2}`n" -f $SIPgroup.IE.('ListenPort_' + $i), $ProtocolLookup.Get_Item($SIPgroup.IE.('Protocol_' + $i)), $TlsProfileIDLookup.Get_Item($SIPgroup.IE.('TLSProfileID_' + $i)))
													}
												}
												$SIPListenList = Strip-TrailingCR -DelimitedString $SIPListenList
												$SipSGTable += ,('Port : Protocol : TLS Profile ID',  $SIPListenList, '', $SIPGroupFederationIP)
												$SipSGTable += ,('SPAN', 'Message Manipulation', '', '')
												#Are we doing message manipulation?
												if ($SIPgroup.IE.IngressSPRMessageTableList -eq $null)
												{
													$SipSGTable += ,('Message Manipulation', '<n/a this rls>', '', '')
												}
												else
												{
													if (($SIPgroup.IE.IngressSPRMessageTableList -eq '') -and ($SIPgroup.IE.EgressSPRMessageTableList -eq ''))
													{
														$SipSGTable += ,('Message Manipulation', 'Disabled', '', '')
													}
													else
													{
														$SIPInboundMsgTrnList = ''
														$SIPOutboundMsgTrnList = ''
														if ($SIPgroup.IE.IngressSPRMessageTableList -ne $null)
														{
															$InboundMsgList = ($SIPgroup.IE.IngressSPRMessageTableList).Split(',')
															foreach ($InboundMsgListEntry in $InboundMsgList)
															{
																$SIPInboundMsgTrnList += $SIPMessageRuleLookup.Get_Item($InboundMsgListEntry)
																$SIPInboundMsgTrnList += "`n"
															}
															$SIPInboundMsgTrnList = Strip-TrailingCR -DelimitedString $SIPInboundMsgTrnList
														}
														if ($SIPgroup.IE.EgressSPRMessageTableList -ne $null)
														{
															$OutboundMsgList = ($SIPgroup.IE.EgressSPRMessageTableList).Split(',')
															foreach ($OutboundMsgListEntry in $OutboundMsgList)
															{
																$SIPOutboundMsgTrnList += $SIPMessageRuleLookup.Get_Item($OutboundMsgListEntry)
																$SIPOutboundMsgTrnList += "`n"
															}
															$SIPOutboundMsgTrnList = Strip-TrailingCR -DelimitedString $SIPOutboundMsgTrnList
														}
														$SipSGTable += ,('Message Manipulation', 'Enabled', '', '')
														$SipSGTable += ,('SPAN-L', 'Inbound Message Table List', 'SPAN-R', 'Outbound Message Table List')
														$SipSGTable += ,('Inbound Message Table List', $SIPInboundMsgTrnList, 'Outbound Message Table List', $SIPOutboundMsgTrnList)
													}
												}
												$SGData += , ('SIP', $SIPgroupDescription, '', $SipSGTable)
											}
											else
											{
												$SIPRecTable = @() #null the collection for each table
												#Consolidate all of the host and mask values (old versions had them separated, newer firmware they're all comma-separated together in the one field):
												$SIPGroupFederationIP = ''
												$SIPGroupRemoteHosts = ($SIPgroup.IE.RemoteHosts).Split(',')
												$SIPGroupRemoteMasks = ($SIPgroup.IE.RemoteMasks).Split(',')
												for ($i = 0 ; $i -le ($SIPGroupRemoteHosts.Count-1); $i++)
												{
													if (($SIPGroupRemoteHosts[$i] -eq '') -and ($SIPGroupRemoteMasks[$i] -eq '')) { continue} #Skip a null entry
													if  ($SIPGroupRemoteMasks[$i] -eq '')
													{
														$SIPGroupFederationIP += $SIPGroupRemoteHosts[$i] + " / <n/a>`n"
													}
													else
													{
														$SIPGroupFederationIP += $SIPGroupRemoteHosts[$i] + ' / ' + $SIPGroupRemoteMasks[$i] + "`n"
													}
												}
												$SIPGroupFederationIP = Strip-TrailingCR -DelimitedString $SIPGroupFederationIP
												if ($SIPGroupFederationIP -eq '') { $SIPGroupFederationIP = '-- Table is empty --' }
												$SIPgroupDescription = Fix-NullDescription -TableDescription $SIPgroup.IE.Description -TableValue $SIPgroup.value -TablePrefix 'SIPREC SG #'
												#Now build the table:
												$SIPRecTable += ,('SPAN', 'SIP Recording Table', '' , '')
												$SIPRecTable += ,('Description', $SIPgroup.IE.Description, '' , '')
												$SIPRecTable += ,('Admin State', $EnabledLookup.Get_Item($SIPgroup.IE.customAdminState), '', '')
												$SIPRecTable += ,('SPAN-L', 'SIP Channels and Routing', 'SPAN-R', 'SIP IP Details')
												#From here I split the page into LH and RH columns to deal with all the possible different combinations.
												$SIPRecL1 = @()
												$SIPRecL2 = @()
												$SIPRecR1 = @()
												$SIPRecR2 = @()
												#LH:
												$SIPRecL1 += 'Channels'
												$SIPRecL2 += $SIPgroup.IE.Channels
												$SIPRecL1 += 'SIP profile'
												$SIPRecL2 += $SipProfileIdLookup.Get_Item($SIPgroup.IE.ProfileID)
												
												$SIPRecL1 += 'Recording Server Table'
												$SIPRecL2 += $SIPServerTablesLookup.Get_Item($SIPgroup.IE.ServerClusterId)
												
												$SIPRecL1 += 'Load Balancing'
												$SIPRecL2 += $SipLoadBalancingLookup.Get_Item($SIPgroup.IE.ServerSelection)
												$SIPRecL1 += 'Channel Hunting'
												$SIPRecL2 += $SgHuntMethodLookup.Get_Item($SIPgroup.IE.HuntMethod)
												#RH:
												$SIPRecR1 += 'Signaling/Media Source IP'
												$SIPRecR2 += $PortToIPAddressLookup.Get_Item($SIPgroup.IE.NetInterfaceSignaling)
												$SIPRecR1 += 'Signaling DSCP'
												$SIPRecR2 += Test-ForNull -LookupTable $null -value $SIPgroup.IE.DSCP
												#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array
												if ($SIPRecL1.Count -ge $SIPRecR1.Count)
												{
													$arrayCount = $SIPRecL1.Count
												}
												else
												{
													$arrayCount = $SIPRecR1.Count
												}
												for ($i = 0; $i -lt $arrayCount; $i++)
												{
													if (($SIPRecL1[$i] -eq ''   ) -and ($SIPRecL2[$i] -eq ''   ) -and ($SIPRecR1[$i] -eq ''   ) -and ($SIPRecR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
													if (($SIPRecL1[$i] -eq $null) -and ($SIPRecL2[$i] -eq $null) -and ($SIPRecR1[$i] -eq $null) -and ($SIPRecR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
													$SIPRecTable += ,($SIPRecL1[$i], $SIPRecL2[$i], $SIPRecR1[$i], $SIPRecR2[$i])
												}
												$SIPRecTable += ,('SPAN-L','Listen Ports', 'SPAN-R', 'Federated IP/FQDN')
												#Build the rows here - Listen Ports first
												$SIPListenList = ''
												for ($i = 1; $i -le 6; $i++)
												{
													if ($SIPgroup.IE.('ListenPort_' + $i) -ne '0') #We have a valid entry.
													{
															$SIPListenList += ("{0} : {1} : {2}`n" -f $SIPgroup.IE.('ListenPort_' + $i), $ProtocolLookup.Get_Item($SIPgroup.IE.('Protocol_' + $i)), $TlsProfileIDLookup.Get_Item($SIPgroup.IE.('TLSProfileID_' + $i)))
													}
												}
												$SIPListenList = Strip-TrailingCR -DelimitedString $SIPListenList
												$SIPRecTable += ,('Port : Protocol : TLS Profile ID',  $SIPListenList, '', $SIPGroupFederationIP)
												$SIPRecTable += ,('SPAN', 'Message Manipulation', '', '')
												#Are we doing message manipulation?
												if ($SIPgroup.IE.IngressSPRMessageTableList -eq $null)
												{
													$SIPRecTable += ,('Message Manipulation', '<n/a this rls>', '', '')
												}
												else
												{
													if (($SIPgroup.IE.IngressSPRMessageTableList -eq '') -and ($SIPgroup.IE.EgressSPRMessageTableList -eq ''))
													{
														$SIPRecTable += ,('Message Manipulation', 'Disabled', '', '')
													}
													else
													{
														$SIPInboundMsgTrnList = ''
														$SIPOutboundMsgTrnList = ''
														if ($SIPgroup.IE.IngressSPRMessageTableList -ne $null)
														{
															$InboundMsgList = ($SIPgroup.IE.IngressSPRMessageTableList).Split(',')
															foreach ($InboundMsgListEntry in $InboundMsgList)
															{
																$SIPInboundMsgTrnList += $SIPMessageRuleLookup.Get_Item($InboundMsgListEntry)
																$SIPInboundMsgTrnList += "`n"
															}
															$SIPInboundMsgTrnList = Strip-TrailingCR -DelimitedString $SIPInboundMsgTrnList
														}
														if ($SIPgroup.IE.EgressSPRMessageTableList -ne $null)
														{
															$OutboundMsgList = ($SIPgroup.IE.EgressSPRMessageTableList).Split(',')
															foreach ($OutboundMsgListEntry in $OutboundMsgList)
															{
																$SIPOutboundMsgTrnList += $SIPMessageRuleLookup.Get_Item($OutboundMsgListEntry)
																$SIPOutboundMsgTrnList += "`n"
															}
															$SIPOutboundMsgTrnList = Strip-TrailingCR -DelimitedString $SIPOutboundMsgTrnList
														}
														$SIPRecTable += ,('Message Manipulation', 'Enabled', '', '')
														$SIPRecTable += ,('SPAN-L', 'Inbound Message Table List', 'SPAN-R', 'Outbound Message Table List')
														$SIPRecTable += ,('Inbound Message Table List', $SIPInboundMsgTrnList, 'Outbound Message Table List', $SIPOutboundMsgTrnList)
													}
												}
												$AllSIPRecData += ,('SIP Recording', $SIPgroupDescription, '', $SIPRecTable)
											}
										}
									}
								}
								# ---- CAS Signaling Groups ----------
								if ($signalinggroup.name -eq 'CAS')
								{
									$casgroups = $signalinggroup.GetElementsByTagName('ID')
									if ($CASgroups.Count -ne 0)
									{
										ForEach ($CASgroup in $CASgroups)
										{
											if ($CASgroup.IE.classname -eq $null) { continue } # Empty / deleted entry
											if ($CASgroup.IE.classname -eq 'CAS_SG_CFG_IE')
											{
												$CasSGTable = @() #null the collection for each table
												$CASgroupDescription = Fix-NullDescription -TableDescription $CASgroup.IE.Description -TableValue $CASgroup.value -TablePrefix 'SG #'
												#The Profile names are stored in separate elements, but as they're all mutually exclusive I want to combine them in one column in the table:
												if ($CASgroup.IE.CasLoopStartFxsProfileId -ne '0')
												{
													$CasProfileName = $CASSignalingProfileLookup.Get_Item($CASgroup.IE.CasLoopStartFxsProfileId)
													$CasType = 'FXS'
												}
												if ($CASgroup.IE.CasLoopStartFxoProfileId -ne '0')
												{
													$CasProfileName = $CASSignalingProfileLookup.Get_Item($CASgroup.IE.CasLoopStartFxoProfileId)
													$CasType = 'FXO'
												}
												if ($CASgroup.IE.CasEnMProfileId -ne '0')
												{
													$CasProfileName = $CASSignalingProfileLookup.Get_Item($CASgroup.IE.CasEnMProfileId)
													$CasType = 'E&M'
												}
												if ($CASgroup.IE.CasR2ProfileId -ne '0')
												{
													$CasProfileName = $CASSignalingProfileLookup.Get_Item($CASgroup.IE.CasR2ProfileId)
													$CasType = 'R2'
												}
												$CasSGTable += ,('SPAN', 'CAS Signaling Group', '' , '')
												$CasSGTable += ,('Description', $CASgroupDescription, '' , '')
												$CasSGTable += ,('Line Type', $CASSgLineTypeLookup.Get_Item($CASgroup.IE.CasLineType), '' , '')
												if ($CASgroup.IE.customAdminState -eq $null)
												{
													$CasSGTable += ,('Admin State', $EnabledLookup.Get_Item($CASgroup.IE.Enabled), '', '')
												}
												else
												{
													$CasSGTable += ,('Admin State', $EnabledLookup.Get_Item($CASgroup.IE.customAdminState), '', '')
												}
												$CasSGTable += ,('Admin State', $EnabledLookup.Get_Item($CASgroup.IE.Enabled), '', '')
												$CasSGTable += ,('SPAN-L', 'Channels and Routing', 'SPAN-R', 'CAS Protocol')
												#Now prepare the table, building the L & R columns as independent arrays, then consolidating them together ready for Word.
												$CasSgL1 = @()
												$CasSgL2 = @()
												$CasSgR1 = @()
												$CasSgR2 = @()
												#LH columns first:
												switch ($CASgroup.IE.CasSgDirection)
												{
													'0' #Inbound
													{
														$CasSgL1 += 'Direction'
														$CasSgL2 += 'Inbound'
														$CasSgL1 += 'Tone Table'
														$CasSgL2 += $ToneTableLookup.Get_Item($CASgroup.IE.ToneTableID)
														$CasSgL1 += 'Action Set Table'
														$CasSgL2 += $ActionSetLookup.Get_Item($CASgroup.IE.ActionSetTableID)
														$CasSgL1 += 'Call Routing Table'
														$CasSgL2 += $CallRoutingTableLookup.Get_Item($CASgroup.IE.RouteTableID)
													}
													'1' #Outbound
													{
														$CasSgL1 += 'Direction'
														$CasSgL2 += 'Outbound'
														$CasSgL1 += 'Channel Hunting'
														$CasSgL2 += $SgHuntMethodLookup.Get_Item($CASgroup.IE.CasSgHuntMethod)
														$CasSgL1 += 'Tone Table'
														$CasSgL2 += $ToneTableLookup.Get_Item($CASgroup.IE.ToneTableID)
													}
													'2' #Bidirectional
													{
														$CasSgL1 += 'Direction'
														$CasSgL2 += 'Bidirectional'
														$CasSgL1 += 'Channel Hunting'
														$CasSgL2 += $SgHuntMethodLookup.Get_Item($CASgroup.IE.CasSgHuntMethod)
														$CasSgL1 += 'Tone Table'
														$CasSgL2 += $ToneTableLookup.Get_Item($CASgroup.IE.ToneTableID)
														$CasSgL1 += 'Action Set Table'
														$CasSgL2 += $ActionSetLookup.Get_Item($CASgroup.IE.ActionSetTableID)
														$CasSgL1 += 'Call Routing Table'
														$CasSgL2 += $CallRoutingTableLookup.Get_Item($CASgroup.IE.RouteTableID)
													}
												}
												$CasSgL1 += 'No Channel Available Override'
												$CasSgL2 += $Q850DescriptionLookup.Get_Item($CASgroup.IE.NoChannelAvailableId)
												$CasSgL1 += 'Call Setup Response Timer'
												$CasSgL2 += ($CASgroup.IE.TimerSanitySetup / 1000)
												#Now the RH side - and out handline depends upon whether the line type is Analog or Digital, hence the apparent repetition.
												$CasSgR1 += 'CAS Signaling Profile'
												$CasSgR2 += ('({0}) {1}' -f ($CasType), ($CasProfileName))
												$CASgroupCallerIdPrivacySignaling = Test-ForNull -LookupTable $null -value $CASgroup.IE.CallerIdPrivacySignaling
												switch ($CASgroup.IE.CasLineType)
												{
													'0' #analog
													{
														switch ($CasType)
														{
															'FXS'
															{
																$CasSgR1 += 'Supplementary Services Profile'
																$CasSgR2 += $CASSupplementaryProfileLookup.Get_Item($CASgroup.IE.CasSupplementaryProfileId)
																$CasSgR1 += 'Caller ID Type'
																$CasSgR2 += $CasSgCallerIDTypeLookup.Get_Item($CASgroup.IE.CallerIDType)
																switch ($CASgroup.IE.CallerIDType)
																{
																	'0'
																	{
																		$CasSgR1 += 'Play Ringback'
																		$CasSgR2 += $CasRingbackLookup.Get_Item($CASgroup.IE.CasSgRingbackPlayPolicy)
																	}
																	{($_ -eq '2') -or ($_ -eq '3') -or ($_ -eq '7')} # The DTMF options
																	{
																		$CasSgR1 += 'Caller ID Start Delimiter'
																		$CasSgR2 += $CasSgDTMFCallerIDDelimiterLookup.Get_Item($CASgroup.IE.CasSgDTMFCallerIDStartDelimiter)
																		$CasSgR1 += 'Caller ID End Delimiter'
																		$CasSgR2 += $CasSgDTMFCallerIDDelimiterLookup.Get_Item($CASgroup.IE.CasSgDTMFCallerIDEndDelimiter)
																		$CasSgR1 += 'Play Ringback'
																		$CasSgR2 += $CasRingbackLookup.Get_Item($CASgroup.IE.CasSgRingbackPlayPolicy)
																		$CasSgR1 += 'Caller ID Privacy Signaling'
																		$CasSgR2 += $CASgroupCallerIdPrivacySignaling

																	}
																	default
																	{
																		$CasSgR1 += 'Play Ringback'
																		$CasSgR2 += $CasRingbackLookup.Get_Item($CASgroup.IE.CasSgRingbackPlayPolicy)
																		$CasSgR1 += 'Caller ID Privacy Signaling'
																		$CasSgR2 += $CASgroupCallerIdPrivacySignaling
																	}
																}
																$CasSgR1 += 'Call Forwarding Feature'
																if ($CASgroup.IE.CallForward -eq '1')
																{
																	$CasSgR2 += 'Enable'
																	$CasSgR1 += 'Call Forwarding Activate DTMF'
																	$CasSgR2 += $CASgroup.IE.CallForwardActivate
																	$CasSgR1 += 'Call Forwarding Deactivate DTMF'
																	$CasSgR2 += $CASgroup.IE.CallForwardDeactivate
																}
																else
																{
																	$CasSgR2 += 'Disable'
																}
															}
															'FXO'
															{
																$CasSgR1 += 'Caller ID Type'
																$CasSgR2 += $CasSgCallerIDTypeLookup.Get_Item($CASgroup.IE.CallerIDType)
																switch ($CASgroup.IE.CallerIDType)
																{
																	{($_ -eq '2') -or ($_ -eq '3') -or ($_ -eq '7')} # The DTMF options
																	{
																		$CasSgR1 += 'Caller ID Start Delimiter'
																		$CasSgR2 += $CasSgDTMFCallerIDDelimiterLookup.Get_Item($CASgroup.IE.CasSgDTMFCallerIDStartDelimiter)
																		$CasSgR1 += 'Caller ID End Delimiter'
																		$CasSgR2 += $CasSgDTMFCallerIDDelimiterLookup.Get_Item($CASgroup.IE.CasSgDTMFCallerIDEndDelimiter)
																	}
																}
																$CasSgR1 += 'Play Ringback'
																$CasSgR2 += $CasRingbackLookup.Get_Item($CASgroup.IE.CasSgRingbackPlayPolicy)
															}
														}
													}
													'1' #Digital
													{
														switch ($CasType)
														{
															'FXS'
															{
																$CasSgR1 += 'Supplementary Services Profile'
																$CasSgR2 += $CASSupplementaryProfileLookup.Get_Item($CASgroup.IE.CasSupplementaryProfileId)
																$CasSgR1 += 'Caller ID Type'
																switch ($CASgroup.IE.CallerIDType)
																{
																	'0'
																	{
																		$CasSgR2 += 'Disabled'
																		$CasSgR1 += 'Play Ringback'
																		$CasSgR2 += $CasRingbackLookup.Get_Item($CASgroup.IE.CasSgRingbackPlayPolicy)
																	}
																	'1'
																	{
																		$CasSgR2 += 'FSK'
																		$CasSgR1 += 'Play Ringback'
																		$CasSgR2 += $CasRingbackLookup.Get_Item($CASgroup.IE.CasSgRingbackPlayPolicy)
																		$CasSgR1 += 'Caller ID Privacy Signaling'
																		$CasSgR2 += $CASgroupCallerIdPrivacySignaling
																	}
																	'2'
																	{
																		$CasSgR2 += 'DTMF'
																		$CasSgR1 += 'Caller ID Start Delimiter'
																		$CasSgR2 += $CasSgDTMFCallerIDDelimiterLookup.Get_Item($CASgroup.IE.CasSgDTMFCallerIDStartDelimiter)
																		$CasSgR1 += 'Caller ID End Delimiter'
																		$CasSgR2 += $CasSgDTMFCallerIDDelimiterLookup.Get_Item($CASgroup.IE.CasSgDTMFCallerIDEndDelimiter)
																		$CasSgR1 += 'Play Ringback'
																		$CasSgR2 += $CasRingbackLookup.Get_Item($CASgroup.IE.CasSgRingbackPlayPolicy)
																		$CasSgR1 += 'Caller ID Privacy Signaling'
																		$CasSgR2 += $CASgroupCallerIdPrivacySignaling
																	}
																}
															}
															'FXO'
															{
																$CasSgR1 += 'Caller ID Type'
																switch ($CASgroup.IE.CallerIDType)
																{
																	'0'
																	{
																		$CasSgR2 += 'Disabled'
																	}
																	'1'
																	{
																		$CasSgR2 += 'FSK or DTMF'
																		$CasSgR1 += 'Caller ID Start Delimiter'
																		$CasSgR2 += $CasSgDTMFCallerIDDelimiterLookup.Get_Item($CASgroup.IE.CasSgDTMFCallerIDStartDelimiter)
																		$CasSgR1 += 'Caller ID End Delimiter'
																		$CasSgR2 += $CasSgDTMFCallerIDDelimiterLookup.Get_Item($CASgroup.IE.CasSgDTMFCallerIDEndDelimiter)
																	}
																}
																$CasSgR1 += 'Play Ringback'
																$CasSgR2 += $CasRingbackLookup.Get_Item($CASgroup.IE.CasSgRingbackPlayPolicy)
															}
															{($_ -eq 'E&M') -or ($_ -eq 'R2')}
															{
																$CasSgR1 += 'Supplementary Services Profile'
																$CasSgR2 += $CASSupplementaryProfileLookup.Get_Item($CASgroup.IE.CasSupplementaryProfileId)
																$CasSgR1 += 'Call Information Format'
																if ($CASgroup.IE.CallInformationFormatENM -eq $null)
																{
																	$CasSgR2 += '<n/a this rls>'
																}
																else
																{
																	switch ($CASgroup.IE.CallInformationFormatENM)
																	{
																		'0' { $CasSgR2 += 'Called Party' }
																		'1' { $CasSgR2 += 'Calling-Called' }
																		'2' { $CasSgR2 += 'Called-Calling' }
																		default { $CasSgR2 += '<Unhandled Value>' }
																	}
																	if (($CASgroup.IE.CallInformationFormatENM -eq '1') -or ($CASgroup.IE.CallInformationFormatENM -eq '2'))
																	{
																		$CasSgR1 += 'Start Digit'
																		$CasSgR2 += $CasSgDTMFCallerIDDelimiterLookup.Get_Item($CASgroup.IE.CasSgENMStartDigit)
																		$CasSgR1 += 'End Digit'
																		$CasSgR2 += $CasSgDTMFCallerIDDelimiterLookup.Get_Item($CASgroup.IE.CasSgENMEndDigit)
																		$CasSgR1 += 'Delimiter Digit'
																		$CasSgR2 += $CasSgDTMFCallerIDDelimiterLookup.Get_Item($CASgroup.IE.CasSgENMDelimiterDigit)
																	}
																}
																$CasSgR1 += 'Play Ringback'
																$CasSgR2 += $CasRingbackLookup.Get_Item($CASgroup.IE.CasSgRingbackPlayPolicy)
															}
														}
													}
												}
												#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array
												if ($CASSgL1.Count -ge $CASSgR1.Count)
												{
													$arrayCount = $CASSgL1.Count
												}
												else
												{
													$arrayCount = $CASSgR1.Count
												}
												for ($i = 0; $i -lt $arrayCount; $i++)
												{
													if (($CASSgL1[$i] -eq ''   ) -and ($CASSgL2[$i] -eq ''   ) -and ($CASSgR1[$i] -eq ''   ) -and ($CASSgR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
													if (($CASSgL1[$i] -eq $null) -and ($CASSgL2[$i] -eq $null) -and ($CASSgR1[$i] -eq $null) -and ($CASSgR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
													$CasSGTable += ,($CasSgL1[$i], $CasSgL2[$i], $CasSgR1[$i], $CasSgR2[$i])
												}
												$CasSGTable += ,('SPAN', 'Assigned Channels', '', '')
												#We need to format the CAS channel numbers before they hit the table:
												$CASApplyToChannelList = [regex]::replace($CASgroup.IE.ApplyToChannelList, ',', "`n")
												$CASNumberList		 = [regex]::replace($CASgroup.IE.ChannelOwnNumberList, ',', "`n")
												$CASHotlineList		= [regex]::replace($CASgroup.IE.FXSHotlineNumberList, ',', "`n")
												$CASCallForwardingList = [regex]::replace($CASgroup.IE.FXSCallForwardingNumberList, ',', "`n")
												if ($CASApplyToChannelList -eq '') { $CASApplyToChannelList = '-- Table is empty --' }
												switch ($CasType)
												{
													{($_ -eq 'FXS') -or ($_ -eq 'E&M') }
													{
														$CasSGTable += ,(("PortName`n{0}" -f ($CASApplyToChannelList)), ("Channel Phone Number`n{0}" -f ($CASNumberList)),("FXS Hotline Number`n{0}" -f ($CASHotlineList)) , ("Call Forwarding Number`n{0}" -f ($CASCallForwardingList)))
													}
													'FXO'
													{
														$CasSGTable += ,(("PortName`n{0}" -f ($CASApplyToChannelList)), ("Channel Phone Number`n{0}" -f ($CASNumberList)), '', '')
													}
												}
											}
											$SGData += , ('CAS', $CASgroupDescription, '', $CasSGTable)
										}
									}
								}
							}
						}
					}
					#---------- LINKED SIGNALING GROUPS (added in 6.1.4) ------------------------
					'SGPeers'
					{
						$SGPeers = $node.GetElementsByTagName('ID')
						$LinkedSGData = @()
						if ($SGPeers.Count -ne 0)
						{
							ForEach ($SGPeer in $SGPeers)
							{
								if ($SGPeer.IE.classname -eq $null) { continue } # Empty / deleted entry
								if ($SGPeer.IE.classname -eq 'SG_PEERS')
								{
									$SGPeerCollection = @() #null the collection for each table
									$SGPeerTitle = Fix-NullDescription -TableDescription $SGPeer.IE.Description -TableValue $SGPeer.value -TablePrefix 'Linked SG Table #'
									$SGPeerSIPGrpList = ''
									$SGPeerISDNGrpList = ''
									if ($SGPeer.IE.SIPSGList -ne $null)
									{
										$SGPeerList = ($SGPeer.IE.SIPSGList).Split(',')
										foreach ($SGPeerEntry in $SGPeerList)
										{
											$SGPeerSIPGrpList += $SgTableLookup.Get_Item($SGPeerEntry)
											$SGPeerSIPGrpList += "`n"
										}
										$SGPeerSIPGrpList = Strip-TrailingCR -DelimitedString $SGPeerSIPGrpList
									}
									if ($SGPeer.IE.ISDNSGList -ne $null)
									{
										$SGPeerISDNList = ($SGPeer.IE.ISDNSGList).Split(',')
										foreach ($SGPeerISDNListEntry in $SGPeerISDNList)
										{
											$SGPeerISDNGrpList += $SgTableLookup.Get_Item($SGPeerISDNListEntry)
											$SGPeerISDNGrpList += "`n"
										}
										$SGPeerISDNGrpList = Strip-TrailingCR -DelimitedString $SGPeerISDNGrpList
									}
									$SGPeerCollection += ,('SPAN-L', 'Linked Signaling Groups', '', '')
									$SGPeerCollection += ,('Description', $SGPeerTitle, '', '')
									$SGPeerCollection += ,('SIP Signaling Group', $SGPeerSIPGrpList, '', '')
									$SGPeerCollection += ,('ISDN Signaling Groups', $SGPeerISDNGrpList, '', '')
									$LinkedSGData += ,($SGPeerTitle, '', '', $SGPeerCollection)
								}
							}
						}
					}


					#---------- TRANSLATIONS ------------------------
					'Translation'
					{
						$Translations = $node.GetElementsByTagName('ID')
						$TranslationData = @()
						$TranslationColumnTitles = @('Entry', 'Enabled', 'Description', 'Input Field', 'Input Field Value', 'Output Field',  'Output Field Value', 'Action')
						if ($Translations.Count -ne 0)
						{
							ForEach ($Translation in $Translations)
							{
								$TranslationTableLookup = Fix-NullAndDuplicateDescriptions -XmlObjectList $Translations -TableListClassName 'CR_TRANSLATION_TABLE_LIST_CONFIG' -TablePrefix 'Translation Table #'
								$TranslationTableLookup.Add('0','None')	# Moved in v6.0 from the central hashtable declarations
								if ($Translation.IE.classname -eq $null) { continue } # Empty / deleted entry
								if ($Translation.IE.classname -eq 'CR_TRANSLATION_TABLE_LIST_CONFIG')
								{
									$TranslationCollection = @() #null the collection for each table
									$translates = $Translation.GetElementsByTagName('ID')
									ForEach ($translate in $translates)
									{
										if ($translate.IE.classname -eq $null) { continue } # Empty / deleted entry
										#We need to turn a handful of numerics back to text:
										#Pull out the special ones, and then default any that remain (as the value there will be literal)
										$InputFieldString = ''
										switch ($translate.IE.InputField)
										{
											'1'	 { $InputFieldString = $NumberTypeLookup.Get_Item($translate.IE.InputFieldValue)}
											'2'	 { $InputFieldString = $NumberingPlanLookup.Get_Item($translate.IE.InputFieldValue)}
											'5'	 { $InputFieldString = $NumberingPlanLookup.Get_Item($translate.IE.InputFieldValue)}
											'6'	 { $InputFieldString = $CallingNumberPresentationLookup.Get_Item($translate.IE.InputFieldValue)}
											'12'	{ $InputFieldString = $NumberingPlanLookup.Get_Item($translate.IE.InputFieldValue)}
											'13'	{ $InputFieldString = $TransferCapabilityLookup.Get_Item($translate.IE.InputFieldValue)}
											default { $InputFieldString = $translate.IE.InputFieldValue}
										}
										$OutputFieldString = ''
										switch ($translate.IE.OutputField)
										{
											'1'	 { $OutputFieldString = $NumberTypeLookup.Get_Item($translate.IE.OutputFieldValue)}
											'2'	 { $OutputFieldString = $NumberingPlanLookup.Get_Item($translate.IE.OutputFieldValue)}
											'5'	 { $OutputFieldString = $NumberingPlanLookup.Get_Item($translate.IE.OutputFieldValue)}
											'6'	 { $OutputFieldString = $CallingNumberPresentationLookup.Get_Item($translate.IE.OutputFieldValue)}
											'12'	{ $OutputFieldString = $NumberingPlanLookup.Get_Item($translate.IE.OutputFieldValue)}
											'13'	{ $OutputFieldString = $TransferCapabilityLookup.Get_Item($translate.IE.OutputFieldValue)}
											default { $OutputFieldString = $translate.IE.OutputFieldValue}
										}
										if ($translate.IE.Description -eq '')
										{
											$TranslationDescription = ('Entry ID {0}' -f $translate.value)
										}
										else
										{
											$TranslationDescription = $translate.IE.Description
										}
										$TranslationValues = @($translate.value, $EnabledLookup.Get_Item($translate.IE.Enabled), $TranslationDescription, $InputFieldLookup.Get_Item($translate.IE.InputField), $InputFieldString, $OutputFieldLookup.Get_Item($translate.IE.OutputField), $OutputFieldString, $TranslateActionLookup.Get_Item($translate.IE.Action))
										$TranslationCollection += , $TranslationValues
									}
									#Now we need to rearrange the entries in the table into their correct sequence:
									if (($Translation.IE.Sequence -ne '') -and ($Translation.IE.Sequence -ne $null))
									{
										$sequence = Decode-Sequence -EncodedSequence $Translation.IE.Sequence -byteCount 2
										$TranslationTemp =@()
										foreach ($reindex in $sequence)
										{
											foreach ($tableRow in $TranslationCollection)
											{
												if ($tableRow[0] -eq $reindex) # (0 is the 'Index' value)
												{
													$TranslationTemp += ,$TableRow
												}
											}
										}
										$TranslationCollection = $TranslationTemp
									}
									$TranslationData += ,($TranslationTableLookup.Get_Item($translation.value), '', $TranslationColumnTitles, $TranslationCollection)
								}
							}
						}
					}

					#---------- TRANSFORMATIONS ------------------------
					'Transformation'
					{
						$transformations = $node.GetElementsByTagName('ID')
						$TransformationData = @()
						$TransformationColumnTitles = @('Entry', 'Enabled', 'Description', 'Input Field', 'Input Field Value', 'Output Field',  'Output Field Value', 'Match Type')
						if ($transformations.Count -ne 0)
						{
							$TransformationTableLookup = Fix-NullAndDuplicateDescriptions -XmlObjectList $transformations -TableListClassName 'CR_TRANSFORMATION_TABLE_LIST_CONFIG' -TablePrefix 'Transformation Table #'
							$TransformationTableLookup.Add('0','None')	# Moved in v6.0 from the central hashtable declarations
							ForEach ($transformation in $transformations)
							{
								if ($transformation.IE.classname -eq $null) { continue } # Empty / deleted entry
								if ($transformation.IE.classname -eq 'CR_TRANSFORMATION_TABLE_LIST_CONFIG')
								{
									$TransformationCollection = @() #null the collection for each table
									$transforms = $transformation.GetElementsByTagName('ID')
									ForEach ($transform in $transforms)
									{
										if ($transform.IE.classname -eq $null) { continue } # Empty / deleted entry
										#We need to turn a handful of numerics back to text:
										#Pull out the special ones, and then default any that remain (as the value there will be literal)
										$InputFieldString = ''
										$MatchTypeString = $MatchTypeLookupLong.Get_Item($transform.IE.MatchType)
										switch ($transform.IE.InputField)
										{
											'1'	 { $InputFieldString = $NumberTypeLookup.Get_Item($transform.IE.InputFieldValue)}
											'2'	 { $InputFieldString = $NumberingPlanLookup.Get_Item($transform.IE.InputFieldValue)}
											'4'	 { $InputFieldString = $NumberTypeLookup.Get_Item($transform.IE.InputFieldValue)}
											'5'	 { $InputFieldString = $NumberingPlanLookup.Get_Item($transform.IE.InputFieldValue)}
											'6'	 { $InputFieldString = $CallingNumberPresentationLookup.Get_Item($transform.IE.InputFieldValue)}
											'7'	 { $InputFieldString = $CallingNumberScreeningLookup.Get_Item($transform.IE.InputFieldValue)}
											'11'	{ $InputFieldString = $NumberTypeLookup.Get_Item($transform.IE.InputFieldValue)}
											'12'	{ $InputFieldString = $NumberingPlanLookup.Get_Item($transform.IE.InputFieldValue)}
											'13'	{ $InputFieldString = $TransferCapabilityLookup.Get_Item($transform.IE.InputFieldValue)}
											'34'	{ if ($transform.IE.InputFieldValue -eq '(.*)')
														{ $InputFieldString = 'All Trunk Groups' } else
														{ $InputFieldString = $transform.IE.InputFieldValue }}
											'36'	{ $MatchTypeString = 'Not Applicable'; $InputFieldString = $transform.IE.InputFieldValue }
											'37'	{ $MatchTypeString = 'Not Applicable'; $InputFieldString = $transform.IE.InputFieldValue }
											default { $InputFieldString = $transform.IE.InputFieldValue }
										}
										$OutputFieldString = ''
										switch ($transform.IE.OutputField)
										{
											'1'	 { $OutputFieldString = $NumberTypeLookup.Get_Item($transform.IE.OutputFieldValue)}
											'2'	 { $OutputFieldString = $NumberingPlanLookup.Get_Item($transform.IE.OutputFieldValue)}
											'4'	 { $OutputFieldString = $NumberTypeLookup.Get_Item($transform.IE.OutputFieldValue)}
											'5'	 { $OutputFieldString = $NumberingPlanLookup.Get_Item($transform.IE.OutputFieldValue)}
											'6'	 { $OutputFieldString = $CallingNumberPresentationLookup.Get_Item($transform.IE.OutputFieldValue)}
											'7'	 { $OutputFieldString = $CallingNumberScreeningLookup.Get_Item($transform.IE.OutputFieldValue)}
											'11'	{ $OutputFieldString = $NumberTypeLookup.Get_Item($transform.IE.OutputFieldValue)}
											'12'	{ $OutputFieldString = $NumberingPlanLookup.Get_Item($transform.IE.OutputFieldValue)}
											'13'	{ $OutputFieldString = $TransferCapabilityLookup.Get_Item($transform.IE.OutputFieldValue)}
											'34'	{ $OutputFieldString = ''}
											default { $OutputFieldString = $transform.IE.OutputFieldValue }
										}
										if ($transform.IE.Description -eq '')
										{
											$TransformationDescription = ('Entry ID {0}' -f $transform.value)
										}
										else
										{
											$TransformationDescription = $transform.IE.Description
										}
										$TransformationValues = @($transform.value, $EnabledLookup.Get_Item($transform.IE.Enabled), $TransformationDescription, $InputFieldLookup.Get_Item($transform.IE.InputField), $InputFieldString, $OutputFieldLookup.Get_Item($transform.IE.OutputField), $OutputFieldString, $MatchTypeString)
										$TransformationCollection += , $TransformationValues
									}
									#Now we need to rearrange the entries in the table into their correct sequence:
									if (($transformation.IE.Sequence -ne '') -and ($transformation.IE.Sequence -ne $null))
									{
										$sequence = Decode-Sequence -EncodedSequence $transformation.IE.Sequence -byteCount 2
										$transformationTemp = @()
										foreach ($reindex in $sequence)
										{
											foreach ($tableRow in $TransformationCollection)
											{
												if ($tableRow[0] -eq $reindex) # (0 is the 'Index' value)
												{
													$transformationTemp += ,$TableRow
												}
											}
										}
										$TransformationCollection = $TransformationTemp
									}
									$TransformationData += ,("Transformations", $TransformationTableLookup.Get_Item($transformation.value), $TransformationColumnTitles, $TransformationCollection)
								}
							}
						}
						if ($TransformationData.Count -gt 1)
						{
							$TransformationData = $TransformationData | Sort-Object -Property @{Expression={$_[1]}; Ascending=$True}
						}
					}

					#---------- CALL ROUTES ------------
					'CallRouting'
					{
						$callroutes = $node.GetElementsByTagName('ID')
						$RouteColumnTitles = @('Enabled', 'Priority', 'Transformation Table', 'Destination Type', 'First Signaling Group', 'Description', 'Fork Call', 'Entry')
						$RouteData = @()
						if ($callroutes.Count -ne 0)
						{
							ForEach ($callroute in $callroutes)
							{
								if ($callroute.IE.classname -eq $null) {continue }
								if ($callroute.IE.classname -eq 'CR_ROUTE_TABLE_LIST_CONFIG')
								{
									$RouteCollection = @() #null the collection for each table
									$IndexedRouteList = @('') #initialise the indexed table list
									$CallRouteDescription = Fix-NullDescription -TableDescription $callroute.IE.Description -TableValue $callroute.value -TablePrefix 'Call Route Table #'
									$routes = $callroute.GetElementsByTagName('ID')
									ForEach ($route in $routes)
									{
										$RouteEntryTable = @() 	#initialise the table list
										if ($route.IE.classname -eq $null) {continue } # Empty / deleted entry
										if ($route.IE.classname -eq 'CR_ROUTE_TABLE_CONFIG')
										{
											[string] $SgList = ''
											#We run through each table twice. The first time generates the overview (as a H-table) where the sequence is highlighted.
											#The second pass then pulls and builds a separate V-table for each of the individual entries
											#FIRST PASS - this generates the summary H-table
											if ($route.IE.SignalingGroupList -ne '')
											{
												#We need to decode the signalling group(s)
												$sequence = Decode-Sequence -EncodedSequence $route.IE.SignalingGroupList -byteCount 2
												if ($sequence.Count -ne 0)
												{
													foreach ($each_SG in $sequence)
													{
														$SgList += ("{0}`n" -f $SgTableLookup.Get_Item($each_SG.ToString()))
													}
													$SgList = $SgList.Substring(0,$SgList.Length-1) #Strip the trailing CR
													$FirstSg = ('{0}' -f $SgTableLookup.Get_Item($sequence[0].ToString()))
												}
											}
											else
											{
												#It'll be a 'Destination Type = Deny' entry
												$SgList = 'None'
												$FirstSg = 'None'
											}
											if ($route.IE.Description -eq '')
											{
												$routeDescription = ('Entry ID {0}' -f $route.value)
											}
											else
											{
												$routeDescription = $route.IE.Description
											}
											#Just in case there are some REALLY old configs still in production somewhere (prior to the introduction of Transformation Tables)
											if ($route.IE.TranslationTable -eq '0')
											{
												$RouteNumberTxfrmTbl = $TransformationTableLookup.Get_Item($route.IE.TransformationTable)
											}
											else
											{
												$RouteNumberTxfrmTbl = $TranslationTableLookup.Get_Item($route.IE.TranslationTable)
											}
											$CRCallForking = Test-ForNull -LookupTable $YesNoLookup -value $route.IE.CallForked
											$RouteValues = @($EnabledLookup.Get_Item($route.IE.Enabled), $route.IE.RoutePriority, $RouteNumberTxfrmTbl, $CRDestinationTypeLookup.Get_Item($route.IE.DestinationType), $FirstSg, $routeDescription, $CRCallForking, $route.value)
											$RouteCollection += ,$RouteValues
											#SECOND PASS - this generates a V-table for each entry in the table
											$RouteEntryTitle = 'Entry ' + $route.value
											if ($route.IE.Description -ne '')
											{
												$RouteEntryTitle += ' - ' + $route.IE.Description
											}
											$RouteEntryTable += ,('SPAN', $RouteEntryTitle, '', '')
											$RouteEntryTable += ,('SPAN-L', 'Route Details', '', '')
											$RouteEntryTable += ,('Description', $routeDescription, '', '')
											$RouteEntryTable += ,('Admin State', $EnabledLookup.Get_Item($route.IE.Enabled), '', '')
											$RouteEntryTable += ,('Route Priority', $route.IE.RoutePriority, '', '')
											$RouteEntryTable += ,('Call Priority', (Test-ForNull -LookupTable $CRCallPriorityLookup -value $route.IE.CallPriority), '', '')
											$RouteEntryTable += ,('Number/Name Transformation Table', $RouteNumberTxfrmTbl, '', '')
											$RouteEntryTable += ,('Time Of Day Restriction', (Test-ForNull -LookupTable $TODTablesLookup -value $route.IE.TimeOfDay), '', '')
											$RouteEntryTable += ,('SPAN-L', 'Destination Information', '', '')
											$RouteEntryTable += ,('Destination Type', $CRDestinationTypeLookup.Get_Item($route.IE.DestinationType), '', '')
											if ($route.IE.DestinationType -eq '2')
											{
												$RouteEntryTable += ,('Deny Q.850 Cause Code', $Q850DescriptionLookup.Get_Item($route.IE.DenyCauseCode), '', '')
											}
											else
											{
												$RouteEntryTable += ,('Message Translation Table', $MsgTranslationTablesLookup.Get_Item($route.IE.MessageTranslationTable), '', '')
												$RouteEntryTable += ,('Cause Code Reroutes', $RerouteTablesLookup.Get_Item($route.IE.ReRouteTable), '', '')
												$RouteEntryTable += ,('Cancel Others upon Forwarding', $EnabledLookup.Get_Item($route.IE.CancelOthersUponForwarding), '', '')
												if ($LicencedForForking -eq $false)
												{
													$RouteEntryTable += ,('Fork Call', 'Not Licensed', '', '')
												}
												else
												{
													$RouteEntryTable += ,('Fork Call', $CRCallForking, '', '')
												}
												$RouteEntryTable += ,('Destination Signaling Groups', $SgList, '', '')
												if ($route.IE.MaximumCallDuration -eq $null)
												{
													$RouteEntryTable += ,('Enable Maximum Call Duration', '<n/a this rls>', '', '')
												}
												else
												{
													if ($route.IE.MaximumCallDuration -eq '0')
													{
														$RouteEntryTable += ,('Enable Maximum Call Duration', 'Disabled', '', '')
													}
													else
													{
														$RouteEntryTable += ,('Enable Maximum Call Duration', 'Enabled', '', '')
														$RouteEntryTable += ,('Maximum Call Duration', ($route.IE.MaximumCallDuration + ' [1..10080] min'), '', '')
													}
												}
												#From here I split the page into LH and RH columns to deal with all the possible different combinations.
												$RouteEntryTableL1 = @()
												$RouteEntryTableL2 = @()
												$RouteEntryTableR1 = @()
												$RouteEntryTableR2 = @()
												#LH:
												$RouteEntryTableL1 += 'SPAN-L'
												$RouteEntryTableL2 += 'Media'
												if ($SWeLite)
												{
													$RouteEntryTableL1 += 'Audio Stream mode'
												}
												else
												{
													$RouteEntryTableL1 += 'Audio/Fax Stream mode'
												}
												$RouteEntryTableL2 += Test-ForNull -LookupTable $CRMediaModeLookup -value $route.IE.MediaMode
												$RouteEntryTableL1 += 'Video/Application Stream mode'
												$RouteEntryTableL2 += Test-ForNull -LookupTable $EnabledLookup -value $route.IE.VideoMediaMode
												if ($SWeLite)
												{
													switch ($route.IE.MediaMode)
													{
														{(($_ -eq '1') -or ($_ -eq '2') -or ($_ -eq '3'))}
														{
															$RouteEntryTableL1 += 'Proxy SRTP Handling'
															if ($route.IE.ProxyHandlingOption -eq '0')
															{
																$RouteEntryTableL2 += 'Relay'
															}
															else
															{
																$RouteEntryTableL2 += 'Local'
															}
														}
													}
												}
												switch ($route.IE.MediaMode)
												{
													{($_ -eq '1') -or ($_ -eq '4') -or ($_ -eq '5')}
													{
														#We don't show Transcoding or the Media List if Media Mode is Proxy1), Direct(4) or Disabled(5)
													}
													default
													{
														$RouteEntryTableL1 += 'Media Transcoding'
														if (($LicencedForTranscoding -eq $false) -and (!$SWeLite))
														{
															$RouteEntryTableL2 += 'Not Licensed'
														}
														else
														{
															$RouteEntryTableL2 += $EnabledLookup.Get_Item($route.IE.MediaTranscoding)
														}
														$RouteEntryTableL1 += 'Media List'
														$RouteEntryTableL2 += $MediaListProfileLookup.Get_Item($route.IE.MediaSelection)
													}
												}
												#RH:
												$RouteEntryTableR1 += 'SPAN-R'
												$RouteEntryTableR2 += 'Quality of Service'
												$RouteEntryTableR1 += 'Quality Metrics Number of Calls'
												$RouteEntryTableR2 += $route.IE.QualityMetricCalls
												$RouteEntryTableR1 += 'Quality Metrics Time Before Retry'
												$RouteEntryTableR2 += $route.IE.QualityMetricTime
												$RouteEntryTableR1 += 'Min. ASR Threshold'
												$RouteEntryTableR2 += $route.IE.QualityMinASRThreshold
												$RouteEntryTableR1 += 'Enable Min MOS Threshold'
												if ($route.IE.QualityMinLQMOSThreshold -eq $null)
												{
													$RouteEntryTableR2 += '<n/a this rls>'
												}
												else
												{
													if ($route.IE.QualityMinLQMOSThreshold -eq '0')
													{
														$RouteEntryTableR2 += 'Disabled'
													}
													else
													{
														$RouteEntryTableR2 += 'Enabled'
														$RouteEntryTableR1 += 'Min. MOS Score'
														$RouteEntryTableR2 += ($route.IE.QualityMinLQMOSThreshold + ' [1.0..5.0]')
													}
												}
												$RouteEntryTableR1 += 'Enable Max. R/T Delay'
												if ($route.IE.QualityMaxRoundTripDelayThreshold -eq '0')
												{
													$RouteEntryTableR2 += 'Disabled'
												}
												else
												{
													$RouteEntryTableR2 += 'Enabled'
													$RouteEntryTableR1 += 'Max. R/T Delay'
													$RouteEntryTableR2 += $route.IE.QualityMaxRoundTripDelayThreshold
												}
												$RouteEntryTableR1 += 'Enable Max. Jitter'
												if ($route.IE.QualityMaxJitterThreshold -eq '0')
												{
													$RouteEntryTableR2 += 'Disabled'
												}
												else
												{
													$RouteEntryTableR2 += 'Enabled'
													$RouteEntryTableR1 += 'Max. Jitter'
													$RouteEntryTableR2 += $route.IE.QualityMaxJitterThreshold
												}
												#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array
												if ($RouteEntryTableL1.Count -ge $RouteEntryTableR1.Count)
												{
													$arrayCount = $RouteEntryTableL1.Count
												}
												else
												{
													$arrayCount = $RouteEntryTableR1.Count
												}
												for ($i = 0; $i -lt $arrayCount; $i++)
												{
													if (($RouteEntryTableL1[$i] -eq ''   ) -and ($RouteEntryTableL2[$i] -eq ''   ) -and ($RouteEntryTableR1[$i] -eq ''   ) -and ($RouteEntryTableR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
													if (($RouteEntryTableL1[$i] -eq $null) -and ($RouteEntryTableL2[$i] -eq $null) -and ($RouteEntryTableR1[$i] -eq $null) -and ($RouteEntryTableR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
													$RouteEntryTable += ,($RouteEntryTableL1[$i], $RouteEntryTableL2[$i], $RouteEntryTableR1[$i], $RouteEntryTableR2[$i])
												}
											}
											#Add sufficient empty values into the array so that we can then poke the value in at the appropriate level
											# (there might be only 3 values, but they might be entries 1, 2 & 5 after the user has deleted some)
											while (($IndexedRouteList.Count -1) -lt $route.Value) { $IndexedRouteList += '' }
											$IndexedRouteList[$route.Value] = $RouteEntryTable #Stick each table into an indexed array so we can re-order below & display in the correct sequence
											$RouteEntryTable = @() #Initialise the table for next time. This is needed here in case an empty table follows - it ensures the below code realises and handles correctly
										}
									}
									#Now we need to rearrange the entries in the overview table into their correct sequence, AND re-order the free-standing tables underneath the same.
									$OrderedRouteList = @()
									if (($CallRoute.IE.Sequence -ne '') -and ($CallRoute.IE.Sequence -ne $null)) #If it exists AND it's not blank
									{
										$sequence = Decode-Sequence -EncodedSequence $CallRoute.IE.Sequence -byteCount 2
										$CallRouteTemp =@()
										foreach ($reindex in $sequence)
										{
											foreach ($tableRow in $RouteCollection)
											{
												if ($tableRow[7] -eq $reindex) # (7 is the 'Index' value)
												{
													$CallRouteTemp += ,$TableRow
												}
											}
											$OrderedRouteList += ,($IndexedRouteList[$reindex])
										}
										$RouteCollection = $CallRouteTemp
									}
									else
									{
										#No sequence? Then we only have one value in the $IndexedRouteList array - which also equals $RouteEntryTable.
										$OrderedRouteList = $IndexedRouteList
									}
									#This saves the overview/heading table:
									$RouteData += ,('Call Routing Table', $CallRouteDescription, $RouteColumnTitles, $RouteCollection)
									#This loop saves the underlying tables (if there are any):
									foreach ($OneTable in $OrderedRouteList)
									{
										if ($OneTable  -eq '')
										{
											# It's empty. (This will be the null entry we initialised the array with
											continue
										}
										$RouteData += ,('Call Routing Table', '', '', $OneTable)
									}
								}
							}
						}
					}

					#---------- TIME OF DAY TABLES ------------
					'TimeOfDay'
					{
						$TODTables = $node.GetElementsByTagName('ID')
						$TODColumnTitles = @('Enabled', 'Description', 'Entry')
						$TODData = @()
						if ($TODTables.Count -ne 0)
						{
							ForEach ($TODTable in $TODTables)
							{
								if ($TODTable.IE.classname -eq $null) {continue } # Empty / deleted entry
								if ($TODTable.IE.classname -eq 'TIME_OF_DAY_TABLE')
								{
									$TODEntryCollection = @() #null the collection for each table
									$TODEntryList = @() #null the list of TOD entries for each table
									$TODTableDescription = Fix-NullDescription -TableDescription $TODTable.IE.Description -TableValue $TODTable.value -TablePrefix 'Time of Day Table #'
									$TODEntries = $TODTable.GetElementsByTagName('ID')
									ForEach ($TODEntry in $TODEntries)
									{
										$TODEntryTable = @() 	#initialise the table list
										if ($TODEntry.IE.classname -eq $null) {continue } # Empty / deleted entry
										if ($TODEntry.IE.classname -eq 'TIME_OF_DAY_ENTRY')
										{
											#We run through each table twice. The first time generates the overview (as a H-table) where the sequence is highlighted.
											#The second pass then pulls and builds a separate V-table for each of the individual entries
											#FIRST PASS - this generates the summary H-table
											if ($TODEntry.IE.Description -eq '')
											{
												$TODEntryDescription = ('Entry ID {0}' -f $TODEntry.value)
											}
											else
											{
												$TODEntryDescription = $TODEntry.IE.Description
											}
											$TODEntryValues = @($EnabledLookup.Get_Item($TODEntry.IE.Enabled), $TODEntryDescription, $TODEntry.value)
											$TODEntryCollection += ,$TODEntryValues
											#SECOND PASS - this generates a V-table for each entry in the table
											$TODEntryEntryTitle = 'Entry ' + $TODEntry.value
											if ($TODEntry.IE.Description -ne '')
											{
												$TODEntryEntryTitle += ' - ' + $TODEntry.IE.Description
											}
											$TODEntryTable += ,('SPAN', $TODEntryEntryTitle, '', '')
											$TODEntryTable += ,('Description', $TODEntryDescription, '', '')
											$TODEntryTable += ,('Admin State', $EnabledLookup.Get_Item($TODEntry.IE.Enabled), '', '')
											$TODEntryTable += ,('Start Time', $TODEntry.IE.StartTime, '', '')
											$TODEntryTable += ,('Stop Time', $TODEntry.IE.StopTime, '', '')
											$TODDaysofWeek = ($TODEntry.IE.DaysOfWeek).Split(',')
											$TODEntryDays = ''
											foreach ($DayofWeek in $TODDaysofWeek)
											{
												$TODEntryDays += $CRTODLookup.Get_Item($DayofWeek)
												$TODEntryDays += "`n"
											}
											$TODEntryDays = Strip-TrailingCR -DelimitedString $TODEntryDays
											$TODEntryTable += ,('Days of Week', $TODEntryDays, '', '')
											$TODEntryList += ,($TODEntryTable)
											$TODEntryTable = @() #Initialise the table for next time. This is needed here in case an empty table follows - it ensures the below code realises and handles correctly
										}
									}
									#This saves the overview/heading table:
									$TODData += ,('Time of Day Table', $TODTableDescription, $TODColumnTitles, $TODEntryCollection)
									#This loop saves the underlying tables (if there are any):
									foreach ($OneTable in $TODEntryList)
									{
										if ($OneTable  -eq '')
										{
											# It's empty.
											continue
										}
										$TODData += ,('Time of Day Table', '', '', $OneTable)
									}
								}
							}
						}
					}


					#---------- ACTION CONFIG ------------
					'ActionConfig'
					{
						$ActionConfigs = $node.GetElementsByTagName('ID')
						$ActionConfigData = @()
						$ActionConfigColumnTitles = @('Description', 'Action', 'Action Param 1', 'Action Param 2', 'Action Param 3', 'Action Param 4')
						$ActionConfigCollection = @() #initialise the collection
						$ActionCollection = @()		  #initialise the collection
						if ($ActionConfigs.Count -ne 0)
						{
							ForEach ($ActionConfig in $ActionConfigs)
							{
								if ($ActionConfig.IE.classname -eq $null) { continue } # Empty / deleted entry
								if ($ActionConfig.IE.classname -eq 'CR_ACTION_CONFIG')
								{
									$ActionSetConfigLookup.Add($ActionConfig.value, $ActionConfig.IE.Description)
									#We need to turn Action Parameter 1 back to text if the Action = 3 / Release call:
									$ActionConfigActionParameter1 = ''
									$ActionConfigActionParameter2 = $ActionConfig.IE.ActionParameter2 #Will add 'ms' suffix if Action=6
									$ActionConfigActionParameter3 = $ActionConfig.IE.ActionParameter3
									$ActionConfigActionParameter4 = $ActionConfig.IE.ActionParameter4
									switch ($ActionConfig.IE.Action)
									{
										'0'	 {$ActionConfigActionParameter1 = $CallRoutingTableLookup.Get_Item($ActionConfig.IE.ActionParameter1)}
										'3'	 {$ActionConfigActionParameter1 = 'Cause Code ' + $Q850DescriptionLookup.Get_Item($ActionConfig.IE.ActionParameter1)}
										'5'	 {$ActionConfigActionParameter1 = $CallRoutingTableLookup.Get_Item($ActionConfig.IE.ActionParameter1); $ActionConfigActionParameter2 = $ActionConfig.IE.ActionParameter2 + ' ms'}
										'6'	 {$ActionConfigActionParameter1 = $ActionConfig.IE.ActionParameter1 + ' ms'}
										'7'	 {$ActionConfigActionParameter1 = $ActionSetLookup.Get_Item($ActionConfig.IE.ActionParameter1)}
										default {$ActionConfigActionParameter1 = $ActionConfig.IE.ActionParameter1}
									}
									if ($ActionConfigActionParameter1 -eq '') {$ActionConfigActionParameter1 = '<n/a>' }
									if ($ActionConfigActionParameter2 -eq '') {$ActionConfigActionParameter2 = '<n/a>' }
									if ($ActionConfigActionParameter3 -eq '') {$ActionConfigActionParameter3 = '<n/a>' }
									if ($ActionConfigActionParameter4 -eq '') {$ActionConfigActionParameter4 = '<n/a>' }
									$ActionConfigValues = @($ActionConfig.IE.Description, $ActionSetActionLookup.Get_Item($ActionConfig.IE.Action), $ActionConfigActionParameter1, $ActionConfigActionParameter2, $ActionConfigActionParameter3, $ActionConfigActionParameter4)
									$ActionCollection += ,$ActionConfigValues
								}
							}
							$ActionConfigData += ,('Call Actions', 'Action Configuration', $ActionConfigColumnTitles, $ActionCollection)
						}
					}

					#---------- ACTION SETS ------------
					'ActionSets'
					{
						$ActionSets = $node.GetElementsByTagName('ID')
						$ActionSetColumnTitles = @('Entry', 'Description', 'Execute If', 'Transformation Table',  'Action On Success', 'Action On Failure')
						if ($ActionSets.Count -ne 0)
						{
							ForEach ($ActionSet in $ActionSets)
							{
								if ($ActionSet.IE.classname -eq $null) { continue } # Empty / deleted entry
								if ($ActionSet.IE.classname -eq 'CR_ACTION_SET_TABLE_LIST_CONFIG')
								{
									$ActionSetCollection = @() #null the collection for each table
									$ActionSetTableDescription = Fix-NullDescription -TableDescription $ActionSet.IE.Description -TableValue $ActionSet.value -TablePrefix 'Action Set Table #'
									$ActionSetEntries = $ActionSet.GetElementsByTagName('ID')
									ForEach ($ActionSetEntry in $ActionSetEntries)
									{
										if ($ActionSetEntry.IE.classname -eq $null) { continue } #Empty / deleted transformation table entry
										if ($ActionSetEntry.IE.classname -eq 'CR_ACTION_SET_TABLE_CONFIG')
										{
											$ActionSetValues = @( $ActionSetEntry.value, $ActionSetEntry.IE.Description, $ActionSetExecutionLookup.Get_Item($ActionSetEntry.IE.ExecuteIF), $TransformationTableLookup.Get_Item($ActionSetEntry.IE.TransformationTable), $ActionSetConfigLookup.Get_Item($ActionSetEntry.IE.ActionOnSuccess), $ActionSetConfigLookup.Get_Item($ActionSetEntry.IE.ActionOnFailure))
											$ActionSetCollection += , $ActionSetValues
										}
									}
									#Now we need to rearrange the entries in the table into their correct sequence:
									if (($ActionSet.IE.Sequence -ne '') -and ($ActionSet.IE.Sequence -ne $null))
									{
										$sequence = Decode-Sequence -EncodedSequence $ActionSet.IE.Sequence -byteCount 2
										$ActionSetTemp =@()
										foreach ($reindex in $sequence)
										{
											foreach ($tableRow in $ActionSetCollection)
											{
												if ($tableRow[0] -eq $reindex) # (0 is the 'Index' value)
												{
													$ActionSetTemp += ,$TableRow
												}
											}
										}
										$ActionSetCollection = $ActionSetTemp
									}
									$ActionConfigData += ,('Call Actions', ('Action Sets - ' + $ActionSetTableDescription), $ActionSetColumnTitles, $ActionSetCollection)
								}
							}
						}
					}

					#---------- Media  ------------
					'Media'
					{
						$systemgroups = $node.GetElementsByTagName('Token')
						ForEach ($systemgroup in $systemgroups)
						{
							# ---- MediaListProfiles ----------
							if ($systemgroup.name -eq 'MediaListProfiles')
							{
								$MediaListData = @()
								$MediaListProfiles = $systemgroup.GetElementsByTagName('ID')
								if ($MediaListProfiles.Count -ne 0)
								{
									ForEach ($MediaListProfile in $MediaListProfiles)
									{
										if ($MediaListProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($MediaListProfile.IE.classname -eq 'MEDIALIST_PROFILE_IE')
										{
											$MediaListTable = @()
											#Like Call Routing Tables, we need to decode the Sequence FIRST, to add the list of media profiles into the table:
											if (($MediaListProfile.IE.VoiceFaxProfileID -ne '') -and ($MediaListProfile.IE.VoiceFaxProfileID -ne $null))
											{
												$sequence = Decode-Sequence -EncodedSequence $MediaListProfile.IE.VoiceFaxProfileID -byteCount 2
												[string] $MediaProfilesList = ''
												if ($sequence.Count -ne 0)
												{
													foreach ($each_List in $sequence)
													{
														$MediaProfilesList += ("{0}`n" -f $VoiceFaxProfilesLookup.Get_Item($each_List.ToString()))
													}
													$MediaProfilesList = $MediaProfilesList.Substring(0,$MediaProfilesList.Length-1) #Strip the trailing CR
												}
											}
											$MediaListTable += ,('SPAN-L', 'Media List', '', '')
											$MediaListProfileDescription = (Fix-NullDescription -TableDescription $MediaListProfile.IE.Description -TableValue $MediaListProfile.value -TablePrefix 'Media List #')
											$MediaListTable += ,('Description', $MediaListProfileDescription, '', '')
											$MediaListTable += ,('Media Profiles List', $MediaProfilesList, '', '')
											$MediaListTable += ,('SDES-SRTP Profile', $SDESMediaCryptoProfileLookup.Get_Item($MediaListProfile.IE.CryptoProfileID) , '', '')
											$MediaListTable += ,('DTLS-SRTP Profile', (Test-ForNull -LookupTable $DTLSMediaCryptoProfileLookup -value $MediaListProfile.IE.DtlsProfileID), '', '')
											$MediaListTable += ,('Media DSCP', (Test-ForNull -LookupTable $null -value $MediaListProfile.IE.DSCP), '', '')
											if (!$SweLite)
											{
												$MediaListTable += ,('RTCP Mode', $MediaRTCPModeLookup.Get_Item($MediaListProfile.IE.MediaRTCPMode), '', '')
											}
											$MediaListTable += ,('Dead Call Detection', $ReverseEnabledLookup.Get_Item($MediaListProfile.IE.DeadCallDetection), '', '')
											$MediaListTable += ,('Silence Suppression', $ReverseEnabledLookup.Get_Item($MediaListProfile.IE.SilenceSuppression), '', '')
											if ($MediaListProfile.IE.CryptoProfileID -ne '0')
											{
												if ($MediaListProfile.IE.SrtpROC -eq $null)
												{
													$MediaListTable += ,('Reset SRTP Rollover Counter', '<n/a this rls>', '', '')
												}
												else
												{
													if (($SWeLite) -or ($releaseBuild -ge 395))
													{
														#Skip it.
													}
													else
													{
														if ($MediaListProfile.IE.SrtpROC -eq 0)
														{
															$MediaListTable += ,('Reset SRTP Rollover Counter', 'Start of Session (RFC 3711)', '', '')
														}
														else
														{
															$MediaListTable += ,('Reset SRTP Rollover Counter', 'Upon New Remote Session Key', '', '')
														}
													}
												}
											}
											if ($SweLite)
											{
												$MediaListTable += ,('SPAN-L', 'Digit Relay', '', '')
												if ($MediaListProfile.IE.DigitRelayType -eq '1')
												{
													$MediaListTable += ,('Digit (DTMF) Relay Type', 'RFC 2833', '', '')
													$MediaListTable += ,('Digit Relay Payload Type', $MediaListProfile.IE.DigitRelayPayloadType, '', '')
												}
												else
												{
													$MediaListTable += ,('Digit (DTMF) Relay Type', 'As Voice', '', '')
												}
											}
											else
											{
												$MediaListTable += ,('SPAN-L', 'Gain Control', 'SPAN-R', 'Digit Relay')
												if ($MediaListProfile.IE.DigitRelayType -eq '1')
												{
													$MediaListTable += ,('Receive Gain', ($MediaListProfile.IE.ReceiveGain + ' dB'), 'Digit (DTMF) Relay Type', 'RFC 2833')
													$MediaListTable += ,('Transmit Gain', ($MediaListProfile.IE.TransmitGain + ' dB'), 'Digit Relay Payload Type',  $MediaListProfile.IE.DigitRelayPayloadType)
												}
												else
												{
													$MediaListTable += ,('Receive Gain', ($MediaListProfile.IE.ReceiveGain + ' dB'), 'Digit (DTMF) Relay Type', 'As Voice')
													$MediaListTable += ,('Transmit Gain', ($MediaListProfile.IE.TransmitGain + ' dB'), '', '')
												}
												$MediaListTable += ,('SPAN-L', 'Passthrough / Tone Detection', '', '')
												$MediaListTable += ,('Modem Passthrough', $ReverseEnabledLookup.Get_Item($MediaListProfile.IE.ModemRelay), '', '')
												$MediaListTable += ,('Fax Passthrough', $ReverseEnabledLookup.Get_Item($MediaListProfile.IE.FaxRelay), '', '')
												$MediaListTable += ,('CNG Tone Detection', $ReverseEnabledLookup.Get_Item($MediaListProfile.IE.CNGToneDetection), '', '')
												$MediaListTable += ,('Fax Tone Detection', (Test-ForNull -LookupTable $EnabledLookup -value $MediaListProfile.IE.FAXToneDetection), '', '')
												if ($MediaListProfile.IE.DtmfSignalToNoise -eq $null)
												{
													$MediaListTable += ,('DTMF Signal to Noise', '<n/a this rls>', '', '')
												}
												else
												{
													$MediaListTable += ,('DTMF Signal to Noise', ($MediaListProfile.IE.DtmfSignalToNoise + ' [-3..+6] dB'), '', '')
												}
												if ($MediaListProfile.IE.DtmfMinimumLevel -eq $null)
												{
													$MediaListTable += ,('DTMF Minimum Level', '<n/a this rls>', '', '')
												}
												else
												{
													$MediaListTable += ,('DTMF Minimum Level', ($MediaListProfile.IE.DtmfMinimumLevel + ' [-48..-14] dBm0'), '', '')
												}
											}
										}
										$MediaListData += , ('Media List', $MediaListProfileDescription, '', $MediaListTable)
									}
								}
							}
							# ---- MediaCryptoProfiles ----------
							if ($systemgroup.name -eq 'MediaCryptoProfiles')
							{
								$SDESProfilesData = @()
								$MediaCryptoProfiles = $systemgroup.GetElementsByTagName('ID')
								if ($MediaCryptoProfiles.Count -ne 0)
								{
									ForEach ($MediaCryptoProfile in $MediaCryptoProfiles)
									{
										if ($MediaCryptoProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($MediaCryptoProfile.IE.classname -eq 'CRYPTO_PROFILE_IE')
										{
											$MediaCryptoProfileDescription = (Fix-NullDescription -TableDescription $MediaCryptoProfile.IE.Description -TableValue $MediaCryptoProfile.value -TablePrefix 'SDES-SRTP Profile #')
											$MediaCryptoProfileTable = @()
											$MediaCryptoProfileTable += ,('SPAN-L', 'SRTP Config', '', '')
											$MediaCryptoProfileTable += ,('Description', $MediaCryptoProfileDescription, '', '')
											$MediaCryptoProfileTable += ,('Operation Option', ($MediaCryptoOperationLookup.Get_Item($MediaCryptoProfile.IE.OperationOption)), '', '')
											$MediaCryptoProfileTable += ,('Crypto Suite', ($MediaCryptoSuiteLookup.Get_Item($MediaCryptoProfile.IE.CryptoSuite)), '', '')

											$MediaCryptoProfileTable += ,('SPAN-L', 'Master Key', '', '')
											if (!($SWeLite))
											{
												if ($MediaCryptoProfile.IE.MasterKeyLifeValue -eq 0)
												{
													$MediaCryptoProfileTable += ,('Master Key Lifetime', 'Never Expires', '', '')
												}
												else
												{
													$MediaCryptoProfileTable += ,('Master Key Lifetime', 'Set', '', '')
													$MediaCryptoProfileTable += ,('Lifetime Value', ('2^' + $MediaCryptoProfile.IE.MasterKeyLifeValue), '', '')
												}
												if ($MediaCryptoProfile.IE.SessionKeyDerivationRate -eq '0')
												{
													$MediaCryptoProfileTable += ,('Derivation Rate', '0', '', '')
												}
												else
												{
													$MediaCryptoProfileTable += ,('Derivation Rate', ('2^' + $MediaCryptoProfile.IE.SessionKeyDerivationRate), '', '')
												}
											}
											$MediaCryptoProfileTable += ,('Key Identifier Length', $MediaCryptoProfile.IE.MasterKeyIdentifierLength, '', '')

											$SDESProfilesData += , ('SDES-SRTP Profiles', $MediaCryptoProfileDescription, '', $MediaCryptoProfileTable)
										}
									}
								}
							}
							# ---- MediaCryptoProfiles ----------
							if ($systemgroup.name -eq 'DTLSProfiles')
							{
								$DTLSProfilesData = @()
								$DTLSProfiles = $systemgroup.GetElementsByTagName('ID')
								if ($DTLSProfiles.Count -ne 0)
								{
									ForEach ($DTLSProfile in $DTLSProfiles)
									{
										if ($DTLSProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($DTLSProfile.IE.classname -eq 'MSC_CFG_DTLS_PROFILE_IE')
										{
											$DTLSProfileDescription = (Fix-NullDescription -TableDescription $DTLSProfile.IE.Description -TableValue $DTLSProfile.value -TablePrefix 'DTLS-SRTP Profile #')
											$DTLSProfileTable = @()
											$DTLSProfileTable += ,('Description', $DTLSProfileDescription, '', '')
											$DTLSProfileTable += ,('SPAN-L', 'DTLS Parameters', '', '')
											$DTLSProfileTable += ,('SPAN-L', 'Common Attributes', '', '')
											$DTLSVersionList = $DTLSProfile.IE.DTLSVersion
											$DTLSVersionList = 'DTLS ' + $DTLSVersionList #Reinstates the "DTLS " prefix that's not captured in the xml
											$DTLSVersionList = [regex]::replace($DTLSVersionList, ',' , "`nDTLS ")
											$DTLSProfileTable += ,('DTLS Version', $DTLSVersionList, '', '')
											$DTLSProfileTable += ,('Mutual Authentication', $EnabledLookup.Get_Item($DTLSProfile.IE.MutualAuth), '', '')
											$DTLSProfileTable += ,('DTLS Handshake Timer', ($DTLSProfile.IE.HandshakeTimer + ' secs [5..60]'), '', '')
											$DTLSProfileTable += ,('Hash Type', $DTLSHashTypeLookup.Get_Item($DTLSProfile.IE.HashType), '', '')
											switch ($DTLSProfile.IE.Role)
											{
												'1' 	{ $DTLSProfileTable += ,('DTLS Role When Answerer', 'Active', '', '') }
												'2' 	{ $DTLSProfileTable += ,('DTLS Role When Answerer', 'Passive', '', '') }
												default { $DTLSProfileTable += ,('DTLS Role When Answerer', '<Unknown>', '', '') }
											}
											$DTLSProfileTable += ,('SPAN-L', 'Client Attributes', '', '')
											$ClientCipherList = ($DTLSProfile.IE.ClientCipherSequence).Split(',')
											$OrderedClientCipherList = ''
											foreach ($ClientCipherEntry in $ClientCipherList)
											{
												$OrderedClientCipherList += $TlsClientCipherLookupV4.Get_Item($ClientCipherEntry)
												$OrderedClientCipherList += "`n"
											}
											$OrderedClientCipherList = Strip-TrailingCR -DelimitedString $OrderedClientCipherList
											$DTLSProfileTable += ,('Client Cipher List', $OrderedClientCipherList, '', '')
											if ($DTLSProfile.IE.MutualAuth -eq '0')
											{
												$DTLSProfileTable += ,('Verify Peer Server Certificate', $EnabledLookup.Get_Item($DTLSProfile.IE.VerifyPeer), '', '')
											}
											$DTLSProfileTable += ,('SPAN-L', 'Server Attribute', '', '')
											$DTLSProfileTable += ,('Cookie Exchange', $EnabledLookup.Get_Item($DTLSProfile.IE.CookieExchange), '', '')
											$DTLSProfileTable += ,('SPAN-L', 'SRTP Parameters', '', '')
											$CryptoSuiteList = ($DTLSProfile.IE.CryptoSuiteSequence).Split(',')
											$OrderedCryptoSuiteList = ''
											foreach ($CryptoSuiteEntry in $CryptoSuiteList)
											{
												$OrderedCryptoSuiteList += $DTLSCryptoSuiteLookup.Get_Item($CryptoSuiteEntry)
												$OrderedCryptoSuiteList += "`n"
											}
											$OrderedCryptoSuiteList = Strip-TrailingCR -DelimitedString $OrderedCryptoSuiteList
											$DTLSProfileTable += ,('Crypto Suite Sequence', $OrderedCryptoSuiteList, '', '')
											$DTLSProfileTable += ,('SPAN-L', 'Master Key', '', '')
											$DTLSProfileTable += ,('Key Identifier Length', $DTLSProfile.IE.MasterKeyIdentifierLength, '', '')
											$DTLSProfilesData += , ('DTLS-SRTP Profiles', $DTLSProfileDescription, '', $DTLSProfileTable)
										}
									}
								}
							}
							# ---- FaxProfiles ----------
							if ($systemgroup.name -eq 'FaxProfiles')
							{
								$MediaFaxData = @()
								$FaxProfiles = $systemgroup.GetElementsByTagName('ID')
								if ($FaxProfiles.Count -ne 0)
								{
									ForEach ($FaxProfile in $FaxProfiles)
									{
										if ($FaxProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($FaxProfile.IE.classname -eq 'FAX_CODEC_PROFILE_IE')
										{
											$FaxProfileDescription = (Fix-NullDescription -TableDescription $FaxProfile.IE.Description -TableValue $FaxProfile.value -TablePrefix 'Fax Profile #')
											$FaxProfileTable = @()
											$FaxProfileTable += ,('SPAN-L', 'Fax Codec Configuration', '', '')
											$FaxProfileTable += ,('Description', $FaxProfileDescription, '', '')
											$FaxProfileTable += ,('Codec', 'T.38 Fax', '', '')
											$FaxProfileTable += ,('Maximum Rate', $FaxProfile.IE.FaxRateInBitsPerSecond, '', '')
											$FaxProfileTable += ,('Signaling Packet Redundancy', $FaxProfile.IE.SignalingPacketRedundancy, '', '')
											$FaxProfileTable += ,('Payload Packet Redundancy', $FaxProfile.IE.PayloadPacketRedundancy, '', '')
											$FaxProfileTable += ,('Error Correction Mode', $ReverseEnabledLookup.Get_Item($FaxProfile.IE.ErrorCorrectionMode), '', '')
											$FaxProfileTable += ,('Training Confirmation Procedure', $FaxTrainingConfirmationLookup.Get_Item($FaxProfile.IE.TrainingConfirmation), '', '')
											$FaxProfileTable += ,('Fallback to Passthrough', $ReverseEnabledLookup.Get_Item($FaxProfile.IE.FaxRelayFallback), '', '')
											$FaxProfileTable += ,('Super G3 to G3 Fallback', (Test-ForNull -LookupTable $EnabledLookup -value $FaxProfile.IE.SG3ToG3Fallback), '', '')

											$MediaFaxData += , ('Media Profiles', $FaxProfileDescription, '', $FaxProfileTable)
										}
									}
								}
							}
							# ---- VoiceProfiles ----------
							if ($systemgroup.name -eq 'VoiceProfiles')
							{
								$MediaProfilesData = @()
								$VoiceProfiles = $systemgroup.GetElementsByTagName('ID')
								if ($VoiceProfiles.Count -ne 0)
								{
									ForEach ($VoiceProfile in $VoiceProfiles)
									{
										if ($VoiceProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($VoiceProfile.IE.classname -eq 'VOICE_CODEC_PROFILE_IE')
										{
											$VoiceProfileDescription = (Fix-NullDescription -TableDescription $VoiceProfile.IE.Description -TableValue $VoiceProfile.value -TablePrefix 'Voice Profile #')
											$VoiceProfileTable = @()
											$VoiceProfileTable += ,('SPAN-L', 'Voice Codec Configuration', '', '')
											$VoiceProfileTable += ,('Description', $VoiceProfileDescription, '', '')
											$VoiceProfileTable += ,('Codec', $MediaTypeLookup.Get_Item($VoiceProfile.IE.MediaType), '', '')
											if ($VoiceProfile.IE.MediaType -eq 21)
											{
												$VoiceProfileTable += ,('Bandwidth', $CodecBandwidthLookup.Get_Item($VoiceProfile.IE.OpusBandwidthSampleRate), '', '')
											}
											if ($VoiceProfile.IE.VoiceRateInBitsPerSecond -ne 0)
											{
												$VoiceProfileTable += ,('Rate', ($VoiceProfile.IE.VoiceRateInBitsPerSecond + ' b/s'), '', '')
											}
											$VoiceProfileTable += ,('Payload Size', ($VoiceProfile.IE.PTimeInMilliSeconds + ' ms'), '', '')
											if (($VoiceProfile.IE.MediaType -eq 5) -or ($VoiceProfile.IE.MediaType -eq 20) -or ($VoiceProfile.IE.MediaType -eq 21)) #We only show Payload Type for G.726, Opus & SILK
											{
												$VoiceProfileTable += ,('Payload Type', $VoiceProfile.IE.PayloadType, '', '')
											}
											if ($VoiceProfile.IE.MediaType -eq 20) #These are unique to Opus
											{
												$VoiceBitRate = "CBR"
												if ($VoiceProfile.IE.VoiceModeBiteRate -eq 1) { $VoiceBitRate = "VBR" }
												$VoiceProfileTable += ,('Voice Bit Rate', $VoiceBitRate, '', '')
											}
											if (($VoiceProfile.IE.MediaType -eq 20) -or ($VoiceProfile.IE.MediaType -eq 21)) #These are unique to Opus & SILK
											{
												$VoiceProfileTable += ,('Use FEC', $TrueFalseLookup.Get_Item($VoiceProfile.IE.UseFEC), '', '')
												$VoiceProfileTable += ,('Use DTX', $TrueFalseLookup.Get_Item($VoiceProfile.IE.UseDTX), '', '')
											}
											if ($VoiceProfile.IE.MediaType -eq 21)
											{
												$VoiceProfileTable += ,('Complexity Level', $VoiceProfile.IE.Complexity, '', '')
											}
											$MediaProfilesData += , ('Media Profiles', $VoiceProfileDescription, '', $VoiceProfileTable)
										}
									}
								}
							}
						}
					}

					#---------- TONE TABLES ------------------------
					'ToneTables'
					{
						$ToneTables = $node.GetElementsByTagName('ID')
						$ToneTableData = @()
						$ToneTableColumnTitles = @('Tone Profile Type', 'Frequency 1', 'Amplitude 1', 'Frequency 2',  'Amplitude 2', 'Cadence on','Cadence off', 'Cadence 2 on','Cadence 2 off')
						if ($ToneTables.Count -ne 0)
						{
							ForEach ($ToneTable in $ToneTables)
							{
								if ($ToneTable.IE.classname -eq $null) { continue } # Empty / deleted entry
								if ($ToneTable.IE.classname -eq 'CONFIG_DESCRIPTION_IE')
								{
									$ToneCollection = @() #null the collection for each table
									$ToneTableTitle = Fix-NullDescription -TableDescription $ToneTable.IE.Description -TableValue $ToneTable.value -TablePrefix 'Tone Table #'
									$tones = $ToneTable.GetElementsByTagName('ID')
									ForEach ($tone in $tones)
									{
										if ($tone.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($SweLite)
										{
											if (($tone.IE.ToneProfileType -eq '1' ) -or ($tone.IE.ToneProfileType -eq '4'))
											{
												#The SweLite only displays Ringback and Congestion tones
											}
											else
											{
												continue
											}
										}
										if (($tone.IE.Cadence2on -eq '0') -and ($tone.IE.Cadence2off -eq '0'))
										{
											$ToneValues = @($ToneProfileTypeLookup.Get_Item($tone.IE.ToneProfileType), $tone.IE.Frequency1, $tone.IE.Amplitude1, $tone.IE.Frequency2, $tone.IE.Amplitude2, $tone.IE.Cadenceon, $tone.IE.Cadenceoff, '<n/a>', '<n/a>')
										}
										else
										{
											$ToneValues = @($ToneProfileTypeLookup.Get_Item($tone.IE.ToneProfileType), $tone.IE.Frequency1, $tone.IE.Amplitude1, $tone.IE.Frequency2, $tone.IE.Amplitude2, $tone.IE.Cadenceon, $tone.IE.Cadenceoff, $tone.IE.Cadence2on, $tone.IE.Cadence2off)
										}
										$ToneCollection += , $ToneValues
									}
								$ToneTableData += ,($ToneTableTitle, '', $ToneTableColumnTitles, $ToneCollection)
								}
							}
						}
					}

					#---------- TELEPHONY MAPPING TABLES ------------------------
					'TelephonyMappingTables'
					{
						$TMTData = @()
						$TMTGroups = $node.GetElementsByTagName('Token')
						ForEach ($TMTGroup in $TMTGroups)
						{
							# ---- Reroute Tables  ----------
							if ($TMTGroup.name -eq 'RerouteTable')
							{
								$RerouteTables = $TMTGroup.GetElementsByTagName('ID')
								$RerouteTablesColumnTitles = @('Description', 'Q.850 Cause Codes')
								$RerouteTableCollection = @()
								if ($RerouteTables.Count -ne 0)
								{
									ForEach ($RerouteTable in $RerouteTables)
									{
										if ($RerouteTable.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($RerouteTable.IE.classname -eq 'CC_REROUTE_CFG_IE')
										{
											$CauseCodeList = '' #null the list for each table
											#NB: There's actually no 'sequencing' of entries going on here. Sonus has stored the re-route codes in an encoded manner,
											# using the same mechanism as the sequence, hence this seemingly out of place call:
											$sequence = Decode-Sequence -EncodedSequence $RerouteTable.IE.CauseCodes -byteCount 1
											foreach ($CauseCode in $sequence)
											{
												$CauseCodeList += $Q850DescriptionLookup.Get_Item($CauseCode.ToString()) + "`n"
											}
											$CauseCodeList = Strip-TrailingCR -DelimitedString $CauseCodeList
											$RerouteTableHeading = Fix-NullDescription -TableDescription $RerouteTable.IE.Description -TableValue $RerouteTable.value -TablePrefix 'Cause Code Reroute Table #'
											$RerouteTableObject = @($RerouteTableHeading, $CauseCodeList)
											$RerouteTableCollection += , $RerouteTableObject
										}
									}
									$TMTData += ,('Cause Code Reroute Table', '', $RerouteTablesColumnTitles, $RerouteTableCollection)
								}
							}

							# ---- SIP to Q850 ----------
							if ($TMTGroup.name -eq 'SipToQ850')
							{
								$SIPToQ850Groups = $TMTGroup.GetElementsByTagName('ID')
								$SIPToQ850GroupsColumnTitles = @('SIP', 'Q.850')
								if ($SIPToQ850Groups.Count -ne 0)
								{
									ForEach ($SIPToQ850Group in $SIPToQ850Groups)
									{
										if ($SIPToQ850Group.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($SIPToQ850Group.IE.classname -eq 'CONFIG_DESCRIPTION_IE')
										{
											$SIPToQ850MapCollection = @() #null the collection for each table
											$SIPToQ850Heading = Fix-NullDescription -TableDescription $SIPToQ850Group.IE.Description -TableValue $SIPToQ850Group.value -TablePrefix 'SIP to Q.850 Cause Code Override Table #'
											$SIPToQ850Maps = $SIPToQ850Group.GetElementsByTagName('ID')
											ForEach ($SIPToQ850Map in $SIPToQ850Maps)
											{
												if ($SIPToQ850Map.IE.classname -eq $null) { continue } # Empty / deleted entry
												$SIPToQ850MapObject = @($SIPDescriptionLookup.Get_Item($SIPToQ850Map.IE.SIP), $Q850DescriptionLookup.Get_Item($SIPToQ850Map.IE.Q850))
												$SIPToQ850MapCollection += , $SIPToQ850MapObject
											}
											$TMTData += ,('SIP To Q.850 Override Tables', $SIPToQ850Heading, $SIPToQ850GroupsColumnTitles, $SIPToQ850MapCollection)
										}
									}
								}
							}

							# ---- Q850 to SIP ----------
							if ($TMTGroup.name -eq 'Q850ToSip')
							{
								$Q850ToSipGroups = $TMTGroup.GetElementsByTagName('ID')
								$Q850ToSipGroupsColumnTitles = @('Q.850', 'SIP')
								if ($Q850ToSipGroups.Count -ne 0)
								{
									ForEach ($Q850ToSipGroup in $Q850ToSipGroups)
									{
										if ($Q850ToSipGroup.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($Q850ToSipGroup.IE.classname -eq 'CONFIG_DESCRIPTION_IE')
										{
											$Q850ToSipMapCollection = @() #null the collection for each table
											$Q850ToSipHeading = Fix-NullDescription -TableDescription $Q850ToSipGroup.IE.Description -TableValue $Q850ToSipGroup.value -TablePrefix 'Q.850 to SIP Cause Code Override Table #'
											$Q850ToSipMaps = $Q850ToSipGroup.GetElementsByTagName('ID')
											ForEach ($Q850ToSipMap in $Q850ToSipMaps)
											{
												if ($Q850ToSipMap.IE.classname -eq $null) { continue } # Empty / deleted entry
												$Q850ToSipMapObject = @($Q850DescriptionLookup.Get_Item($Q850ToSipMap.IE.Q850),$SIPDescriptionLookup.Get_Item($Q850ToSipMap.IE.SIP))
												$Q850ToSipMapCollection += , $Q850ToSipMapObject
											}
											$TMTData += ,('Q.850 To SIP Override Tables', $Q850ToSipHeading, $Q850ToSipGroupsColumnTitles, $Q850ToSipMapCollection)
										}
									}
								}
							}

							# ---- MsgTranslationTable ----------
							if ($TMTGroup.name -eq 'MsgTranslationTable')
							{
								$MsgTranslationTables = $TMTGroup.GetElementsByTagName('ID')
								$MsgTranslationTableColumnTitles = @('Entry', 'Enabled', 'Description')
								if ($MsgTranslationTables.Count -ne 0)
								{
									ForEach ($MsgTranslationTable in $MsgTranslationTables)
									{
										if ($MsgTranslationTable.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($MsgTranslationTable.IE.classname -eq 'MESSAGE_PI_TRANSLATION_TABLE_IE')
										{
											$MsgTranslationCollection = @() #null the collection for each table
											$MTTableList = @('')			#initialise the indexed table list
											$MsgTranslationHeading = Fix-NullDescription -TableDescription $MsgTranslationTable.IE.Description -TableValue $MsgTranslationTable.value -TablePrefix 'Message Translation Table #'
											$MsgTranslationTableEntries = $MsgTranslationTable.GetElementsByTagName('ID')
											ForEach ($MsgTranslationTableEntry in $MsgTranslationTableEntries)
											{
												$MTTable = @() 	#initialise the table list
												if ($MsgTranslationTableEntry.IE.classname -eq $null) { continue } # Empty / deleted entry
												if ($MsgTranslationTableEntry.IE.classname -eq 'MESSAGE_PI_TRANSLATION_IE')
												{
													#We run through each table twice. The first time generates the overview (as a H-table) where the sequence is highlighted.
													#The second pass then pulls and builds a separate V-table for each of the individual entries
													#FIRST PASS - this generates the summary H-table
													$MsgTranslationTableEntryObject = @($MsgTranslationTableEntry.value, $EnabledLookup.Get_Item($MsgTranslationTableEntry.IE.Enabled), $MsgTranslationTableEntry.IE.Description)
													$MsgTranslationCollection += , $MsgTranslationTableEntryObject

													#SECOND PASS - this generates a V-table for each entry in the table
													$MTEntryTitle = 'Entry ' + $MsgTranslationTableEntry.Value
													if ($MsgTranslationTableEntry.IE.Description -ne '')
													{
														$MTEntryTitle += ' - ' + $MsgTranslationTableEntry.IE.Description
													}
													$MTTable += ,('SPAN', $MTEntryTitle, '', '')
													$MTTable += ,('Description', $MsgTranslationTableEntry.IE.Description, '', '')
													$MTTable += ,('Enabled', $YesNoLookup.Get_Item($MsgTranslationTableEntry.IE.Enabled), '', '')
													$MTTable += ,('SPAN-L', 'Incoming Message', 'SPAN-R', 'Outgoing')
													#From here I split the page into LH and RH columns to deal with all the possible different combinations.
													$MTTableL1 = @()
													$MTTableL2 = @()
													$MTTableR1 = @()
													$MTTableR2 = @()
													#LH:
													$MTTableL1 += 'Message Type'
													$MTTableL2 += $MsgXlatMsgTypeLookup.Get_Item($MsgTranslationTableEntry.IE.InMsgType)
													$MTTableL1 += 'IE Type / SDP Presence'
													switch ($MsgTranslationTableEntry.IE.InIEType)
													{
														'0'
														{
															$MTTableL2 += 'Progress Indicator / SDP'
															$MTTableL1 += 'ISDN PI Present'
															if ($MsgTranslationTableEntry.IE.InPIPresent -eq '0')
															{
																$MTTableL2 += 'No'
																#That's it for here
															}
															else
															{
																$MTTableL2 += 'Yes'
																$MTTableL1 += 'ISDN PI Values'
																if ($MsgTranslationTableEntry.IE.InPiValue -eq '0')
																{
																	$MTTableL2 += 'Any'
																	#That's it for here
																}
																else
																{
																	$MTTableL2 += 'Selected'
																	$MTTableL1 += 'Progress Indicators'
																	#'1' = 'Not end to end ISDN'; '2' = 'Destination not ISDN'; '4' = 'Origin not ISDN'; '8' = 'Return to ISDN'; '16' = 'Interworking Encountered'; '32' = 'Inband Information'; '64' = 'Delay Encountered'
																	$TMTPIValue = $MsgTranslationTableEntry.IE.InPiValue #An encoded value of the above elements
																	$TMTPIText = ''
																	if (($TMTPIValue - 64) -ge 0)
																	{
																		$TMTPIValue -= 64
																		$TMTPIText = 'Delay Encountered'
																	}
																	if (($TMTPIValue - 32) -ge 0)
																	{
																		$TMTPIValue -= 32
																		$TMTPIText = 'Inband Information,' + $TMTPIText
																	}
																	if (($TMTPIValue - 16) -ge 0)
																	{
																		$TMTPIValue -= 16
																		$TMTPIText = 'Interworking Encountered,' + $TMTPIText
																	}
																	if (($TMTPIValue - 8) -ge 0)
																	{
																		$TMTPIValue -= 8
																		$TMTPIText = 'Return to ISDN,' + $TMTPIText
																	}
																	if (($TMTPIValue - 4) -ge 0)
																	{
																		$TMTPIValue -= 4
																		$TMTPIText = 'Origin not ISDN,' + $TMTPIText
																	}
																	if (($TMTPIValue - 2) -ge 0)
																	{
																		$TMTPIValue -= 2
																		$TMTPIText = 'Destination not ISDN,' + $TMTPIText
																	}
																	if (($TMTPIValue - 1) -ge 0)
																	{
																		$TMTPIText = 'Not end to end ISDN,' + $TMTPIText
																	}
																	$TMTPIText = [regex]::replace($TMTPIText, ',' , "`n") # Write each attribute on a new line
																	$TMTPIText = Strip-TrailingCR -DelimitedString $TMTPIText
																	$MTTableL2 += $TMTPIText
																}
															}
														}
														'1'
														{
															$MTTableL2 += "Don't Care"
															#That's it for here
														}
														'2'
														{
															$MTTableL2 += 'Facility'
															#That's it for here
														}
													}
													#RH:
													$MTTableR1 += 'Send Message'
													$MTTableR2 += $MsgXlatSendMessageLookup.Get_Item($MsgTranslationTableEntry.IE.SendMessage)
													if ($MsgTranslationTableEntry.IE.SendMessage -eq '1')
													{
														#That's it for this side of the table!
													}
													else
													{
														$MTTableR1 += 'Message Type'
														$MTTableR2 += $MsgXlatMsgTypeLookup.Get_Item($MsgTranslationTableEntry.IE.OutMsgType)
														$MTTableR1 += 'IE / SDP Presence'
														$MTTableR2 += $MsgXlatOutIELookup.Get_Item($MsgTranslationTableEntry.IE.OutIEType)
														$MTTableR1 += 'Media Cut Through'
														$MTTableR2 += $MsgXlatMediaCutThroughTypeLookup.Get_Item($MsgTranslationTableEntry.IE.MediaCutThroughType)
													}
													#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array
													if ($MTTableL1.Count -ge $MTTableR1.Count)
													{
														$arrayCount = $MTTableL1.Count
													}
													else
													{
														$arrayCount = $MTTableR1.Count
													}
													for ($i = 0; $i -lt $arrayCount; $i++)
													{
														if (($MTTableL1[$i] -eq ''   ) -and ($MTTableL2[$i] -eq ''   ) -and ($MTTableR1[$i] -eq ''   ) -and ($MTTableR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
														if (($MTTableL1[$i] -eq $null) -and ($MTTableL2[$i] -eq $null) -and ($MTTableR1[$i] -eq $null) -and ($MTTableR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
														$MTTable += ,($MTTableL1[$i], $MTTableL2[$i], $MTTableR1[$i], $MTTableR2[$i])
													}
													if ($MsgTranslationTableEntry.IE.InIEType -ne '2')
													{
														$MTTable += ,('SPAN-L', 'Media', '', '')
														$MTTable += ,('Early Media Status', $MsgXlatEarlyMediaStatusLookup.Get_Item($MsgTranslationTableEntry.IE.EarlyMediaStatus), '', '')
													}
													#Add sufficient empty values into the array so that we can then poke the value in at the appropriate level
													# (there might be only 3 values, but they might be entries 1, 2 & 5 after the user has deleted some)
													while (($MTTableList.Count -1) -lt $MsgTranslationTableEntry.Value) { $MTTableList += '' }
													$MTTableList[$MsgTranslationTableEntry.Value] = $MTTable #Stick each table into an indexed array so we can re-order below & display in the correct sequence
													$MTTable = @() #Initialise the table for next time. This is needed here in case an empty table follows - it ensures the below code realises and handles correctly
												}
											}
											#Now we need to rearrange the entries in the overview table into their correct sequence, AND re-order the free-standing tables underneath the same.
											$OrderedMTTableList = @()
											if (($MsgTranslationTable.IE.Sequence -ne '') -and ($MsgTranslationTable.IE.Sequence -ne $null))
											{
												$sequence = Decode-Sequence -EncodedSequence $MsgTranslationTable.IE.Sequence -byteCount 1
												$MsgTranslationTemp =@()
												foreach ($reindex in $sequence)
												{
													foreach ($tableRow in $MsgTranslationCollection)
													{
														if ($tableRow[0] -eq $reindex) # (0 is the 'Index' value)
														{
															$MsgTranslationTemp += ,$TableRow
														}
													}
													$OrderedMTTableList += ,($MTTableList[$reindex])
												}
												$MsgTranslationCollection = $MsgTranslationTemp
											}
											else
											{
												#No sequence? Then we only have one value in the $MMTableList array - which also equals $MTTable.
												$OrderedMTTableList = $MTTableList
											}
											#This saves the overview/heading table:
											$TMTData += ,('Message Translations', $MsgTranslationHeading, $MsgTranslationTableColumnTitles, $MsgTranslationCollection)
											#This loop saves the underlying tables (if there are any):
											foreach ($OneTable in $OrderedMTTableList)
											{
												if ($OneTable  -eq '')
												{
													# It's empty. (This will be the null entry we initialised the array with
													continue
												}
												$TMTData += ,('Message Translations', '', '', $OneTable)
											}
										}
									}
								}
							}
						}
					}

					#---------- SIP ------------------------
					'SIP'
					{
						$SIPGroups = $node.GetElementsByTagName('Token')
						ForEach ($SIPGroup in $SIPGroups)
						{
							# ---- SIP Servers ----------
							if ($SIPGroup.name -eq 'SIPServers')
							{
								$AllSipServerData = @()
								$SIPServerGroups = $SIPGroup.GetElementsByTagName('ID')
								$SIPServerGroupsColumnTitles = @('Entry', 'Host / Domain', 'Server Lookup', 'Port', 'Protocol')
								if ($SIPServerGroups.Count -ne 0)
								{
									ForEach ($SIPServerGroup in $SIPServerGroups)
									{
										if ($SIPServerGroup.IE.classname -eq $null) { continue } # Empty / deleted entry
										if (($SIPServerGroup.IE.classname -eq 'CONFIG_DESCRIPTION_IE') -or ($SIPServerGroup.IE.classname -eq 'SIP_CFG_SERVER_TABLE_LIST_IE')) #The latter is new in Rls 3.0
										{
											$SIPServerCollection = @() #null the collection for each table
											$SIPServerTableList = @('')			#initialise the indexed table list
											$SIPServerTableHeading = Fix-NullDescription -TableDescription $SIPServerGroup.IE.Description -TableValue $SIPServerGroup.value -TablePrefix 'SIP Server Table #'
											$SIPServers = $SIPServerGroup.GetElementsByTagName('ID')
											ForEach ($SIPServer in $SIPServers)
											{
												$SipServerTable = @() 	#initialise the table list
												if ($SIPServer.IE.classname -eq $null) { continue } # Empty / deleted entry
												if ($SIPServer.IE.classname -eq 'SIP_CFG_SERVER_IE')
												{
													#We run through each table twice. The first time generates the overview (as a H-table) where the sequence is highlighted.
													#The second pass then pulls and builds a separate V-table for each of the individual entries
													#FIRST PASS - this generates the summary H-table
													if ($SIPServer.IE.ServerLookup -eq $null)
													{
														#If no Lookup option, force back to 'host'
														$SIPServerDestination = $SIPServer.IE.Host
														$SIPServerLookup = 'IP/FQDN'
													}
													else
													{
														if ($SIPServer.IE.ServerLookup -eq '0')
														{
															$SIPServerLookup = 'IP/FQDN'
															$SIPServerDestination = $SIPServer.IE.Host
														}
														else
														{
															$SIPServerLookup = 'DNS SRV'
															$SIPServerDestination =  $SIPServer.IE.DomainName
														}
													}
													$SIPServerObject = @($SIPServer.Value, $SIPServerDestination, $SIPServerLookup, $SIPServer.IE.Port, $ProtocolLookup.Get_Item($SIPServer.IE.Protocol))
													$SIPServerCollection += , $SIPServerObject
													#SECOND PASS - this generates a V-table for each entry in the table
													$SipServerEntryTitle = 'Entry ' + $SIPServer.Value + ' : ' + $SIPServerDestination
													# if ($SIPServerGroup.IE.Description -ne '')
													# {
														# $SipServerEntryTitle += ' - ' + $SIPServerGroup.IE.Description
													# }
													$SipServerTable += ,('SPAN', $SipServerEntryTitle, '', '')
													$SipServerTable += ,('SPAN-L', 'Server Host', 'SPAN-R', 'Transport')
													#From here I split the page into LH and RH columns to deal with all the possible different combinations.
													$SSTableL1 = @()
													$SSTableL2 = @()
													$SSTableR1 = @()
													$SSTableR2 = @()
													#LH:
													$SSTableL1 += 'Server Lookup'
													if (($SIPServer.IE.ServerLookup -eq $null) -or ($SIPServer.IE.ServerLookup -eq '0'))
													{
														$SSTableL2 += 'IP/FQDN'
														$SSTableL1 += 'Priority'
														if ($SIPServer.IE.Priority -eq 0)
														{
															$SSTableL2 += 'None'
														}
														else
														{
															$SSTableL2 += $SIPServer.IE.Priority
														}
														$SSTableL1 += 'Host'
														$SSTableL2 += $SIPServer.IE.Host
														if ($SIPServer.IE.HostIpVersion -eq $null)
														{
															# Don't show 'Host IP Version'
														}
														else
														{
															try
															{
																if ([ipaddress]$SIPServer.IE.Host) {}
																# It's an IP address: Don't show 'Host IP Version'
															}
															catch
															{
																#It's not a valid IP, or it's a hostname
																$SSTableL1 += 'Host IP Version'
																$SSTableL2 += Test-ForNull -LookupTable $IpVersionLookup -value $SIPServer.IE.HostIpVersion
															}
														}
														$SSTableL1 += 'Port'
														$SSTableL2 += $SIPServer.IE.Port
													}
													else
													{
														$SSTableL2 += 'DNS SRV'
														$SSTableL1 += 'Domain Name / FQDN'
														$SSTableL2 += $SIPServer.IE.DomainName
														$SSTableL1 += 'Service Name'
														$SSTableL2 += $SIPServer.IE.ServiceName
													}
													$SSTableL1 += 'Protocol'
													$SSTableL2 += $ProtocolLookup.Get_Item($SIPServer.IE.Protocol)
													if ($SIPServer.IE.Protocol -eq '4')
													{
														$SSTableL1 += 'TLS Profile'
														$SSTableL2 += $TlsProfileIdLookup.Get_Item($SIPServer.IE.TLSProfileID)
													}
													#RH:
													$SSTableR1 += 'Monitor'
													if ($SIPServer.IE.Monitor -eq '0')
													{
														$SSTableR2 += 'None'
													}
													else
													{
														$SSTableR2 += 'SIP Options'
														$SSTableR1 += 'Keep Alive Frequency'
														$SSTableR2 += $SIPServer.IE.KeepAliveFrequency + ' secs'
														$SSTableR1 += 'Recover Frequency'
														$SSTableR2 += $SIPServer.IE.RecoverFrequency + ' secs'
														$SSTableR1 += 'Local Username'
														$SSTableR2 += Test-ForNull -LookupTable $null -value $SIPServer.IE.LocalUserName
														$SSTableR1 += 'Peer Username'
														$SSTableR2 += Test-ForNull -LookupTable $null -value $SIPServer.IE.PeerUserName
													}
													#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array
													if ($SSTableL1.Count -ge $SSTableR1.Count)
													{
														$arrayCount = $SSTableL1.Count
													}
													else
													{
														$arrayCount = $SSTableR1.Count
													}
													for ($i = 0; $i -lt $arrayCount; $i++)
													{
														if (($SSTableL1[$i] -eq ''   ) -and ($SSTableL2[$i] -eq ''   ) -and ($SSTableR1[$i] -eq ''   ) -and ($SSTableR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
														if (($SSTableL1[$i] -eq $null) -and ($SSTableL2[$i] -eq $null) -and ($SSTableR1[$i] -eq $null) -and ($SSTableR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
														$SipServerTable += ,($SSTableL1[$i], $SSTableL2[$i], $SSTableR1[$i], $SSTableR2[$i])
													}
													#Go again for the lower half of the table:
													# *Skip* if v9.0.0+ and this is a SIP Recorder (SipServerType = 1)
													if (($SIPServer.IE.SipServerType -ne $null) -and ($SIPServer.IE.SipServerType -eq 1))
													{
														$SSTableL1 = @()
														$SSTableL2 = @()
														$SSTableR1 = @()
														$SSTableR2 = @()
														#LH:
														$SSTableL1 += 'SPAN-L'
														$SSTableL2 += 'Remote Authorization and Contacts'
														$SSTableL1 += 'Remote Authorization Table'
														$SSTableL2 += $SipCredentialsTableLookup.Get_Item($SIPServer.IE.RemoteAuthorizationTableID)
														if ($SIPServer.IE.ContactRegistrantTableID -eq '0')
														{
															$SSTableL1 += 'Contact Registrant Table'
															$SSTableL2 += 'None'
														}
														else
														{
															$SSTableL1 += 'Contact Registrant Table'
															$SSTableL2 += $SipRegistrationTableLookup.Get_Item($SIPServer.IE.ContactRegistrantTableID)
															$SSTableL1 += 'Clear Remote Registration on Startup'
															$SSTableL2 += Test-ForNull -LookupTable $TrueFalseLookup -value $SIPServer.IE.ClearRemoteRegistrationOnStartup
															$SSTableL1 += 'Contact URI Randomizer'
															$SSTableL2 += Test-ForNull -LookupTable $TrueFalseLookup -value $SIPServer.IE.ContactURIRandomizer
															$SSTableL1 += 'Stagger Registration'
															$SSTableL2 += Test-ForNull -LookupTable $TrueFalseLookup -value $SIPServer.IE.StaggerRegistration
														}
														if ($SIPServer.IE.RemoteAuthorizationTableID -ne '0')
														{
															#Only displays (in 4.1+ ?) if we have a Remote Auth Table set
															$SSTableL1 += 'Retry Non-Stale Nonce'
															$SSTableL2 += Test-ForNull -LookupTable $TrueFalseLookup -value $SIPServer.IE.RetryNonStaleNonce
															$SSTableL1 +=  'Authorization on Refresh'
															$SSTableL2 += Test-ForNull -LookupTable $TrueFalseLookup -value $SIPServer.IE.AuthorizationOnRefresh
														}
														$SSTableL1 += 'Session URI Validation'
														if ($SIPServer.IE.SessionURIValidation -eq $null)
														{
															$SSTableL2 += '<n/a this rls>'
														}
														else
														{
															if ($SIPServer.IE.SessionURIValidation -eq '1')
															{
																$SSTableL2 += 'Strict'
															}
															else
															{
																$SSTableL2 += 'Liberal'
															}
														}
														#RH:
														if ($SIPServer.IE.Protocol -ne '1') #(Do for everything *except* UDP)
														{
															$SSTableR1 += 'SPAN-R'
															$SSTableR2 += 'Connection Reuse'
															$SSTableR1 += 'Reuse'
															if ($SIPServer.IE.ReuseTransport -eq '0')
															{
																$SSTableR2 += 'False'
															}
															else
															{
																$SSTableR2 += 'True'
																$SSTableR1 += 'Sockets'
																$SSTableR2 += $SIPServer.IE.TransportSocket
																$SSTableR1 += 'Reuse Timeout'
																if ($SIPServer.IE.ReuseTimeout -eq '0')
																{
																	$SSTableR2 += 'Forever'
																}
																else
																{
																	$SSTableR2 += 'Limited'
																	$SSTableR1 += 'Timeout Limit'
																	$SSTableR2 += $SIPServer.IE.ReuseTimeout + ' mins'
																}
															}
														}
														#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array
														if ($SSTableL1.Count -ge $SSTableR1.Count)
														{
															$arrayCount = $SSTableL1.Count
														}
														else
														{
															$arrayCount = $SSTableR1.Count
														}
														for ($i = 0; $i -lt $arrayCount; $i++)
														{
															if (($SSTableL1[$i] -eq ''   ) -and ($SSTableL2[$i] -eq ''   ) -and ($SSTableR1[$i] -eq ''   ) -and ($SSTableR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
															if (($SSTableL1[$i] -eq $null) -and ($SSTableL2[$i] -eq $null) -and ($SSTableR1[$i] -eq $null) -and ($SSTableR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
															$SipServerTable += ,($SSTableL1[$i], $SSTableL2[$i], $SSTableR1[$i], $SSTableR2[$i])
														}
													} #End the lower half skip code added in v9.0.0 for SIP Recorders
													#Add sufficient empty values into the array so that we can then poke the value in at the appropriate level
													# (there might be only 3 values, but they might be entries 1, 2 & 5 after the user has deleted some)
													while (($SIPServerTableList.Count -1) -lt $SIPServer.Value) { $SIPServerTableList += '' }
													$SIPServerTableList[$SIPServer.Value] = $SipServerTable #Stick each table into an indexed array so we can re-order below & display in the correct sequence
													$SipServerTable = @() #Initialise the table for next time. This is needed here in case an empty table follows - it ensures the below code realises and handles correctly
												}
											}
											#Now we need to rearrange the entries in the table into their correct sequence:
											#(Sequence here was added in Rls3)
											$OrderedSipServerTableList = @()
											if (($SIPServerGroup.IE.Sequence -ne '') -and ($SIPServerGroup.IE.Sequence -ne $null)) #If it exists AND it's not blank
											{
												$sequence = Decode-Sequence -EncodedSequence $SIPServerGroup.IE.Sequence -byteCount 2
												$SipServerTemp =@()
												foreach ($reindex in $sequence)
												{
													foreach ($tableRow in $SIPServerCollection)
													{
														if ($tableRow[0] -eq $reindex) # (0 is the 'Index' value)
														{
															$SipServerTemp += ,$TableRow
														}
													}
													$OrderedSipServerTableList += ,($SIPServerTableList[$reindex])
												}
												$SIPServerCollection = $SipServerTemp
											}
											else
											{
												#No sequence? Then we only have one value in the $SIPServerTableList array - which also equals $SipServerTable.
												$OrderedSIPServerTableList = $SipServerTableList
											}
											#This saves the overview/heading table:
											$AllSipServerData += ,('SIP Server Tables', $SIPServerTableHeading, $SIPServerGroupsColumnTitles, $SIPServerCollection)
											#This loop saves the underlying tables (if there are any):
											foreach ($OneTable in $OrderedSIPServerTableList)
											{
												if ($OneTable -eq '')
												{
													# It's empty. (This will be the null entry we initialised the array with
													continue
												}
												$AllSipServerData += ,('SIP Server Tables', '', '', $OneTable)
											}
										}
									}
								}
							}

							# ---- SIP Registrars ----------
							if ($SIPGroup.name -eq 'SIPRegistrars')
							{
								$SIPRegistrars = $SIPGroup.GetElementsByTagName('ID')
								$SIPRegistrarsColumnTitles = @('Entry', 'Description', 'Limit Users', 'Number of Users')
								$SIPRegistrarCollection = @()
								$AllSipLocalRegistrarData = @()
								if ($SIPRegistrars.Count -ne 0)
								{
									ForEach ($SIPRegistrar in $SIPRegistrars)
									{
										if ($SIPRegistrar.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($SIPRegistrar.IE.classname -eq 'SIP_CFG_REGISTRAR_IE')
										{
											if ($SIPRegistrar.IE.MaxUsers -eq '0')
											{
												$SIPRegistrarLimitUsers = 'No'
												$SIPRegistrarMaxUsers = 'Unlimited'
											}
											else
											{
												$SIPRegistrarLimitUsers = 'Yes'
												$SIPRegistrarMaxUsers = $SIPRegistrar.IE.MaxUsers
											}
											$SIPRegistrarDescription = Fix-NullDescription -TableDescription $SIPRegistrar.IE.Description -TableValue $SIPRegistrar.value -TablePrefix 'Local Registrar #'
											$SIPRegistrarObject = @($SIPRegistrar.Value, $SIPRegistrarDescription, $SIPRegistrarLimitUsers, $SIPRegistrarMaxUsers)
											$SIPRegistrarCollection += , $SIPRegistrarObject
										}
									}
									$AllSipLocalRegistrarData += ,('Local Registrars', '', $SIPRegistrarsColumnTitles, $SIPRegistrarCollection)
								}
							}

							if ($SIPGroup.name -eq 'SIPRegistrationTables')
							{
								$AllSipContactRegistrantData = @()
								$SIPRegistrationTables = $SIPGroup.GetElementsByTagName('ID')
								$SIPRegistrationTablesColumnTitles = @('Entry', 'Type of Address of Record', 'Address of Record URI', 'Global TTL', 'Failed Registration Retry', 'Contact URI Username', 'TTL', 'Priority (Q)')
								if ($SIPRegistrationTables.Count -ne 0)
								{
									ForEach ($SIPRegistrationTable in $SIPRegistrationTables)
									{
										if ($SIPRegistrationTable.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($SIPRegistrationTable.IE.classname -eq 'CONFIG_DESCRIPTION_IE')
										{
											$SIPRegistrationTableCollection = @() #null the collection for each table
											$SIPRegistrationTableHeading = Fix-NullDescription -TableDescription $SIPRegistrationTable.IE.Description -TableValue $SIPRegistrationTable.value -TablePrefix 'Contact Registrant Table #'
											$SIPRegistrations = $SIPRegistrationTable.GetElementsByTagName('ID')
											ForEach ($SIPRegistration in $SIPRegistrations)
											{
												if ($SIPRegistration.IE.classname -eq $null) { continue } # Empty / deleted entry
												# Consolidate any/all Contact URI Usernames, TTL & Q values into single strings
												$SIPRegoUserNamesList = ''
												$SIPRegoUserTTLList  = ''
												$SIPRegoUserQList  = ''
												for ($i = 1; $i -le 8; $i++)
												{
													$temp = 'TTL' + $i
													if ($SIPRegistration.IE.('URI' + $i) -ne $null)
													{
														#OK, so it exists in the XML, but is it populated?
														if ($SIPRegistration.IE.('URI' + $i) -ne '')
														{
															$SIPRegoUserNamesList += ("{0} `n" -f $SIPRegistration.IE.('URI' + $i))
															if ($SIPRegistration.IE.('TTL' + $i) -ne $null)
															{
																if ($SIPRegistration.IE.('TTL' + $i) -eq '0')
																{
																	$SIPRegoUserTTLList += "Inherited`n"
																}
																else
																{
																	$SIPRegoUserTTLList += ("{0} `n" -f $SIPRegistration.IE.('TTL' + $i))
																}
															}
															if ($SIPRegistration.IE.('Q' + $i) -ne $null)
															{
																$SIPRegoUserQList += ("{0} `n" -f ($SIPRegistration.IE.('Q' + $i)/100))
															}
														}
													}
												}
												$SIPRegoUserNamesList = Strip-TrailingCR -DelimitedString $SIPRegoUserNamesList
												$SIPRegoUserTTLList = Strip-TrailingCR -DelimitedString $SIPRegoUserTTLList
												$SIPRegoUserQList = Strip-TrailingCR -DelimitedString $SIPRegoUserQList
												$SIPRegistrationObject = @($SIPRegistration.Value, $AORTypeLookup.Get_Item($SIPRegistration.IE.AorType), $SIPRegistration.IE.AOR, $SIPRegistration.IE.AorTtl, (Test-ForNull -LookupTable $null -value $SIPRegistration.IE.FailedRegistrationRetryTimer), $SIPRegoUserNamesList, $SIPRegoUserTTLList, $SIPRegoUserQList)
												$SIPRegistrationTableCollection += , $SIPRegistrationObject
											}
											$AllSipContactRegistrantData += ,('Contact Registrant', $SIPRegistrationTableHeading, $SIPRegistrationTablesColumnTitles, $SIPRegistrationTableCollection)
										}
									}
								}
							}

							if ($SIPGroup.name -eq 'SIPAuthorizationTables') #On-screen these  are 'Local Pass-thu Auth Tables' (SIPAuthorisationsTableData)
							{
								$AllSipLocalPassthroughData = @()
								$SIPAuthorisationsTables = $SIPGroup.GetElementsByTagName('ID')
								$SIPAuthorisationsTablesColumnTitles = @('Entry', 'Address of Record Type', 'Address of Record URI', 'User Name')
								$SIPAuthorisationIeMatchRegex = ''
								if ($SIPAuthorisationsTables.Count -ne 0)
								{
									ForEach ($SIPAuthorisationsTable in $SIPAuthorisationsTables)
									{
										if ($SIPAuthorisationsTable.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($SIPAuthorisationsTable.IE.classname -eq 'CONFIG_DESCRIPTION_IE')
										{
											$SIPAuthorisationsTableCollection = @() #null the collection for each table
											$SIPAuthorisationsTableHeading = Fix-NullDescription -TableDescription $SIPAuthorisationsTable.IE.Description -TableValue $SIPAuthorisationsTable.value -TablePrefix 'Local Pass-thu Authorization Table #'
											$SIPAuthorisations = $SIPAuthorisationsTable.GetElementsByTagName('ID')
											ForEach ($SIPAuthorisation in $SIPAuthorisations)
											{
												if ($SIPAuthorisation.IE.classname -eq $null) { continue } # Empty / deleted entry
												$SIPAuthorisationsObject = @($SIPAuthorisation.Value, $AORTypeLookup.Get_Item($SIPAuthorisation.IE.AorType), $SIPAuthorisation.IE.AOR, $SIPAuthorisation.IE.User)
												$SIPAuthorisationsTableCollection += , $SIPAuthorisationsObject
											}
											$AllSipLocalPassthroughData += ,('Local / Pass-through Auth Tables', $SIPAuthorisationsTableHeading, $SIPAuthorisationsTablesColumnTitles, $SIPAuthorisationsTableCollection)
										}
									}
								}
							}

							if ($SIPGroup.name -eq 'SIPCredentialsTables') #On-screen these  are 'Remote Authorization Tables'
							{
								$AllSipRemoteAuthData = @()
								$SIPCredentialsTables = $SIPGroup.GetElementsByTagName('ID')
								$SIPCredentialsTablesColumnTitles = @('Entry', 'Realm', 'Authentication ID', 'From URI User Match', 'Match Regex')
								$SIPCredentialIeMatchRegex = ''
								if ($SIPCredentialsTables.Count -ne 0)
								{
									ForEach ($SIPCredentialsTable in $SIPCredentialsTables)
									{
										if ($SIPCredentialsTable.IE.classname -eq $null) { continue } # Empty / deleted entry
										if (($SIPCredentialsTable.IE.classname -eq 'SIP_CFG_USER_CREDENTIALS_LIST_IE') -or ($SIPCredentialsTable.IE.classname -eq 'CONFIG_DESCRIPTION_IE')) #Changed to the former in Rls 3.
										{
											$SIPCredentialsTableCollection = @() #null the collection for each table
											$SIPCredentialsTableHeading = Fix-NullDescription -TableDescription $SIPCredentialsTable.IE.Description -TableValue $SIPCredentialsTable.value -TablePrefix 'Remote Authorization Table #'
											$SIPCredentials = $SIPCredentialsTable.GetElementsByTagName('ID')
											ForEach ($SIPCredential in $SIPCredentials)
											{
												if ($SIPCredential.IE.classname -eq $null) { continue } # Empty / deleted entry
												if ($SIPCredential.IE.FromURIUserMatch -eq $null)
												{
													# Not in this firmware
														$SIPCredentialsObject = @($SIPCredential.Value, $SIPCredential.IE.Realm, $SIPCredential.IE.User, '<n/a this rls>', '<n/a this rls>')
												}
												else
												{
													if ($SIPCredential.IE.FromURIUserMatch -eq '0')
													{
														# Don't display any value remaining in this cell
														$SIPCredentialIeMatchRegex = 'N/A'
													}
													else
													{
														$SIPCredentialIeMatchRegex = $SIPCredential.IE.MatchRegex
													}
													$SIPCredentialsObject = @($SIPCredential.Value, $SIPCredential.IE.Realm, $SIPCredential.IE.User, $RemoteAuthFromURILookup.Get_Item($SIPCredential.IE.FromURIUserMatch), $SIPCredentialIeMatchRegex)
												}
												$SIPCredentialsTableCollection += , $SIPCredentialsObject
											}
											#Now we need to rearrange the entries in the table into their correct sequence:
											#(Sequence here was added in Rls3, and it's specified as an absolute, rather than encoded, as all previous sequences have been. e.g. it's now '3,1,2')
											if (($SIPCredentialsTable.IE.Sequence -ne '') -and ($SIPCredentialsTable.IE.Sequence -ne $null)) #If it exists AND it's not blank
											{
												$SIPCredentialTemp =@()
												$CredentialSequence = ($SIPCredentialsTable.IE.Sequence).Split(',')
												foreach ($reindex in $CredentialSequence)
												{
													foreach ($tableRow in $SIPCredentialsTableCollection)
													{
														if ($tableRow[0] -eq $reindex) # (0 is the 'Index' value)
														{
															$SIPCredentialTemp += ,$TableRow
														}
													}
												}
												$SIPCredentialsTableCollection = $SIPCredentialTemp
											}
											$AllSipRemoteAuthData += ,('Remote Authorization Tables', $SIPCredentialsTableHeading, $SIPCredentialsTablesColumnTitles, $SIPCredentialsTableCollection)
										}
									}
								}
							}

							# ---- SIP Profiles ----------
							if ($SIPGroup.name -eq 'Profile')
							{
								$SIPProfiles = $SIPGroup.GetElementsByTagName('ID')
								$AllSIPProfileData = @()
								if ($SIPProfiles.Count -ne 0)
								{
									ForEach ($SIPProfile in $SIPProfiles)
									{
										if ($SIPProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($SIPProfile.IE.classname -eq 'SIP_CFG_PROFILE_IE')
										{
											$SipProfilesTable = @()
											#We need to check for a value not present in older firmware:
											$SIPProfileDescription = Fix-NullDescription -TableDescription $SIPProfile.IE.Description -TableValue $SIPProfile.value -TablePrefix 'SIP Profile #'
											$SipProfilesTable += ,('SPAN', 'SIP Profile', '' , '')
											$SipProfilesTable += ,('Description', $SIPProfileDescription, '' , '')
											$SipProfilesTable += ,('SPAN-L', 'Session Timer', 'SPAN-R' , 'MIME Payloads')
											#Now prepare the table, building the L & R columns as independent arrays, then consolidating them together ready for Word.
											$SipProfilesL1 = @()
											$SipProfilesL2 = @()
											$SipProfilesR1 = @()
											$SipProfilesR2 = @()
											#LH columns first:
											$SipProfilesL1 += 'Session Timer'
											if ($SIPProfile.IE.SessionTimer -eq '1')
											{
												$SipProfilesL2 += 'Enable'
												$SipProfilesL1 += 'Refresh method'
												$SipProfilesL2 += Test-ForNull -LookupTable $SipProfileRefreshLookup -value $SIPProfile.IE.RefreshMethod #Removed in/from around 3.1.1b290
												$SipProfilesL1 += 'Minimum Acceptable Timer'
												$SipProfilesL2 += $SIPProfile.IE.SessionTimerMin
												$SipProfilesL1 += 'Offered Session Timer'
												$SipProfilesL2 += $SIPProfile.IE.SessionTimerExp
												$SipProfilesL1 += 'Terminate on Refresh Failure'
												$SipProfilesL2 += Test-ForNull -LookupTable $TrueFalseLookup -value $SIPProfile.IE.TerminateOnRefreshFailure
											}
											else
											{
												$SipProfilesL2 += 'Disable'
											}
											#Now the RH:
											$SipProfilesR1 += 'ELIN Identifier'
											$SipProfilesR2 += Test-ForNull -LookupTable $SipProfileElinIDLookup -value $SIPProfile.IE.ElinIdentifier
											$SipProfilesR1 += 'PIDF-LO Passthrough'
											$SipProfilesR2 += Test-ForNull -LookupTable $EnableLookup -value $SIPProfile.IE.PidfPlPassthru
											$SipProfilesR1 += 'Unknown Subtype Passthrough'
											$SipProfilesR2 += Test-ForNull -LookupTable $EnableLookup -value $SIPProfile.IE.UnknownPlPassthru
											#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array
											if ($SipProfilesL1.Count -ge $SipProfilesR1.Count)
											{
												$arrayCount = $SipProfilesL1.Count
											}
											else
											{
												$arrayCount = $SipProfilesR1.Count
											}
											for ($i = 0; $i -lt $arrayCount; $i++)
											{
												if (($SipProfilesL1[$i] -eq ''   ) -and ($SipProfilesL2[$i] -eq ''   ) -and ($SipProfilesR1[$i] -eq ''   ) -and ($SipProfilesR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
												if (($SipProfilesL1[$i] -eq $null) -and ($SipProfilesL2[$i] -eq $null) -and ($SipProfilesR1[$i] -eq $null) -and ($SipProfilesR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
												$SipProfilesTable += ,($SipProfilesL1[$i], $SipProfilesL2[$i], $SipProfilesR1[$i], $SipProfilesR2[$i])
											}
											$SipProfilesTable += ,('SPAN-L', 'Header Customization', 'SPAN-R' , 'Options Tags')
											#Here we go again: the tables slip out of alignment if session timer is enabled.
											$SipProfilesL1 = @()
											$SipProfilesL2 = @()
											$SipProfilesR1 = @()
											$SipProfilesR2 = @()
											#LH columns first:
											if (($SWeLite) -or ($releaseBuild -ge 353))
											{
												# 'Subscription State Passthrough' vanished from ~v4.0.0, assumed to be build 353.
											}
											else
											{
												$SipProfilesL1 += 'Subscription State Passthrough'
												$SipProfilesL2 += $EnableLookup.Get_Item($SIPProfile.IE.AllowHeader)
											}
											$SipProfilesL1 += 'FQDN in From Header'
											$SipProfilesL2 += Test-ForNull -LookupTable $SipProfileFrmHdrLookup -value $SIPProfile.IE.FQDNinFromHeader
											if (($SIPProfile.IE.FQDNinFromHeader -eq '3') -or ($SIPProfile.IE.FQDNinContactHeader -eq '3'))
											{
												$SipProfilesL1 += 'Static Host FQDN/IP[:port]'
												$SipProfilesL2 += $SIPProfile.IE.StaticHost
											}
											$SipProfilesL1 += 'FQDN in Contact Header'
											$SipProfilesL2 += Test-ForNull -LookupTable $SipProfileCntctHdrLookup -value $SIPProfile.IE.FQDNinContactHeader
											$SipProfilesL1 += 'Send Assert Header'
											$SipProfilesL2 += Test-ForNull -LookupTable $SipProfileAssertHdrLookup -value $SIPProfile.IE.SendAssertHdrAlways
											$SipProfilesL1 += 'Sonus Diagnostics Header'
											$SipProfilesL2 += Test-ForNull -LookupTable $EnableLookup -value $SIPProfile.IE.DiagnosticsHeader
											$SipProfilesL1 += 'Trusted Interface'
											$SipProfilesL2 += Test-ForNull -LookupTable $EnableLookup -value $SIPProfile.IE.TrustedInterface
											if ($SIPProfile.IE.TrustedInterface -eq '1')
											{
												$SipProfilesL1 += 'UA Header'
												$SipProfilesL2 += $SIPProfile.IE.UserAgentHeader
											}
											$SipProfilesL1 += 'Calling Info Source'
											$SipProfilesL2 += Test-ForNull -LookupTable $SipProfileClgInfoSourceLookup -value $SIPProfile.IE.CgNumberNameFromHdr
											$SipProfilesL1 +=  'Diversion Header Selection'
											$SipProfilesL2 += Test-ForNull -LookupTable $SipProfileDivHrdLookup -value $SIPProfile.IE.DiversionSelection
											$SipProfilesL1 +=  'Record Route Header'
											$SipProfilesL2 += Test-ForNull -LookupTable $SipProfileRecordHrdLookup -value $SIPProfile.IE.RecordRouteHdrPref
											#RH columns
											$SipProfilesR1 += '100rel'
											$SipProfilesR2 += $SipProfileOptionsLookup.Get_Item($SIPProfile.IE.Option100Rel)
											$SipProfilesR1 += 'Path'
											$SipProfilesR2 += Test-ForNull -LookupTable $SipProfileOptionsLookup -value $SIPProfile.IE.OptionPath
											if ($SIPProfile.IE.SessionTimer -eq '1')
											{
												$SipProfilesR1 += 'Timer'
												$SipProfilesR2 += $SipProfileOptionsLookup.Get_Item($SIPProfile.IE.OptionTimer)
											}
											$SipProfilesR1 += 'Update'
											$SipProfilesR2 += $SipProfileOptionsLookup.Get_Item($SIPProfile.IE.OptionUpdate)
											#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array (which will always be Left)
											for ($i = 0; $i -lt $SipProfilesL1.Count; $i++)
											{
												if (($SipProfilesL1[$i] -eq ''   ) -and ($SipProfilesL2[$i] -eq ''   ) -and ($SipProfilesR1[$i] -eq ''   ) -and ($SipProfilesR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
												if (($SipProfilesL1[$i] -eq $null) -and ($SipProfilesL2[$i] -eq $null) -and ($SipProfilesR1[$i] -eq $null) -and ($SipProfilesR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
												$SipProfilesTable += ,($SipProfilesL1[$i], $SipProfilesL2[$i], $SipProfilesR1[$i], $SipProfilesR2[$i])
											}
											$SipProfilesTable += ,('SPAN-L', 'Timers', 'SPAN-R' , 'SDP Customization')
											#One more time:
											$SipProfilesL1 = @()
											$SipProfilesL2 = @()
											$SipProfilesR1 = @()
											$SipProfilesR2 = @()
											#LH columns first:
											$SipProfilesL1 += 'Transport Timeout Timer'
											$SipProfilesL2 += Test-ForNull -LookupTable $null -value $SIPProfile.IE.TransportTimeoutTimer
											$SipProfilesL1 += 'Maximum Retransmissions'
											$SipProfilesL2 += Test-ForNull -LookupTable $SipProfileMaxReTxLookup -value $SIPProfile.IE.MaxRetransmits
											$SipProfilesL1 += 'Redundancy Retry Timer'
											$SipProfilesL2 += Test-ForNull -LookupTable $null -value $SIPProfile.IE.RedundancyRetryTimer
											$SipProfilesL1 += 'SPAN-L'
											$SipProfilesL2 += 'RFC Timers'
											$SipProfilesL1 += 'Timer T1'
											$SipProfilesL2 += $SIPProfile.IE.TimerT1
											$SipProfilesL1 += 'Timer T2'
											$SipProfilesL2 += $SIPProfile.IE.TimerT2
											$SipProfilesL1 += 'Timer T4'
											$SipProfilesL2 += $SIPProfile.IE.TimerT4
											$SipProfilesL1 += 'Timer D'
											$SipProfilesL2 += $SIPProfile.IE.TimerD
											$SipProfilesL1 += 'Timer J'
											$SipProfilesL2 += Test-ForNull -LookupTable $null -value $SIPProfile.IE.TimerJ
											#RH columns
											$SipProfilesR1 += 'Send Number of Audio Channels'
											$SipProfilesR2 += Test-ForNull -LookupTable $TrueFalseLookup -value $SIPProfile.IE.SendNumberofAudioChan
											$SipProfilesR1 += 'Connection Info in Media Section'
											$SipProfilesR2 += Test-ForNull -LookupTable $TrueFalseLookup -value $SIPProfile.IE.ConnectionInfoInMediaSection
											$SipProfilesR1 += 'Origin Field Username'
											$SipProfilesR2 += Test-ForNull -LookupTable $null -value $SIPProfile.IE.OriginFieldUserName
											$SipProfilesR1 += 'Session Name'
											$SipProfilesR2 += Test-ForNull -LookupTable $null -value $SIPProfile.IE.SessionName
											$SipProfilesR1 += 'Digit Transmission Preference'
											$SipProfilesR2 += Test-ForNull -LookupTable $SipProfileDigitPrefLookup -value $SIPProfile.IE.DigitPreference
											$SipProfilesR1 += 'SDP Handling Preference'
											$SipProfilesR2 += Test-ForNull -LookupTable $SipProfileSDPHandlingLookup -value $SIPProfile.IE.SDPHandling
											#Reassemble the above back into the correct column appearances for Word. Reconstitute for the length of the larger array (which will always be Left)
											for ($i = 0; $i -lt $SipProfilesL1.Count; $i++)
											{
												if (($SipProfilesL1[$i] -eq ''   ) -and ($SipProfilesL2[$i] -eq ''   ) -and ($SipProfilesR1[$i] -eq ''   ) -and ($SipProfilesR2[$i] -eq ''   )) {continue} #No point writing a totally blank row!
												if (($SipProfilesL1[$i] -eq $null) -and ($SipProfilesL2[$i] -eq $null) -and ($SipProfilesR1[$i] -eq $null) -and ($SipProfilesR2[$i] -eq $null)) {continue} #No point writing a totally blank row!
												$SipProfilesTable += ,($SipProfilesL1[$i], $SipProfilesL2[$i], $SipProfilesR1[$i], $SipProfilesR2[$i])
											}
										}
										$AllSIPProfileData += ,('SIP Profiles', $SIPProfileDescription, '', $SipProfilesTable)
									}
								}
							}

							 # ---- Trunk Groups ----------
							if ($SIPGroup.name -eq 'SIPTrunkGroup')
							{
								$SIPTrunkGroups = $SIPGroup.GetElementsByTagName('ID')
								$AllSIPTrunkGroupData = @()
								if ($SIPTrunkGroups.Count -ne 0)
								{
									ForEach ($SIPTrunkGroup in $SIPTrunkGroups)
									{
										if ($SIPTrunkGroup.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($SIPTrunkGroup.IE.classname -eq 'SIP_CFG_TRUNK_GROUP_IE')
										{
											$SIPTrunkGroupsTable = @()
											$SIPTrunkGroupDescription = Fix-NullDescription -TableDescription $SIPTrunkGroup.IE.Description -TableValue $SIPTrunkGroup.value -TablePrefix 'Trunk Group ID: '
											$SIPTrunkGroupsTable += ,('SPAN-L', 'SIP Trunk Group Table', '' , '')
											$SIPTrunkGroupsTable += ,('Description', $SIPTrunkGroupDescription, '' , '')
											$SIPTrunkGroupsTable += ,('Trunk Group ID', $SIPTrunkGroup.IE.TrunkGroupID, '' , '')
											if ($SIPTrunkGroup.IE.Type -eq '0')
											{
												$SIPTrunkGroupsTable += ,('Trunk Group Type', 'DTG/OTG', '' , '')
											}
											else
											{
												$SIPTrunkGroupsTable += ,('Trunk Group Type', 'TGRP' , '' , '')
												$SIPTrunkGroupsTable += ,('Trunk-Context', $SIPTrunkGroup.IE.TrunkContext, '' , '')
											}
											$SIPTrunkGroupsTable += ,('Include ID in Outbound Calls', $EnabledLookup.Get_Item($SIPTrunkGroup.IE.IncludeIDInOutboundCalls), '' , '')
											$SIPTrunkGroupsTable += ,('Use ID for Routing Inbound Calls', $EnabledLookup.Get_Item($SIPTrunkGroup.IE.UseIDForRoutingInboundCalls), '' , '')
											#We need to decode the signalling group(s)
											$TrunkSgList = ''
											$sequence = Decode-Sequence -EncodedSequence $SIPTrunkGroup.IE.SignalingGroupList -byteCount 2
											if ($sequence.Count -ne 0)
											{
												foreach ($each_SG in $sequence)
												{
													$TrunkSgList += ("{0}`n" -f $SgTableLookup.Get_Item($each_SG.ToString()))
												}
												$TrunkSgList = $TrunkSgList.Substring(0,$TrunkSgList.Length-1) #Strip the trailing CR
											}
											$SIPTrunkGroupsTable += ,('Associated Signaling Groups', $TrunkSgList, '' , '')
										}
										$AllSIPTrunkGroupData += ,('Trunk Groups', $SIPTrunkGroupDescription, '', $SIPTrunkGroupsTable)
									}
								}
							}

							# ---- SIP NAT Qualified Prefix Tables ----------
							if ($SIPGroup.name -eq 'SIPNatPrefixTable')
							{
								$AllSIPNATPrefixData = @()
								$SIPNATPrefixTables = $SIPGroup.GetElementsByTagName('ID')
								$SIPNATPrefixColumnTitles = @('Entry', 'Description', 'IP Address', 'Netmask')
								if ($SIPNATPrefixTables.Count -ne 0)
								{
									ForEach ($SIPNATPrefixTable in $SIPNATPrefixTables)
									{
										if ($SIPNATPrefixTable.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($SIPNATPrefixTable.IE.classname -eq 'SIP_CFG_NAT_PREFIX_TABLE_LIST_IE')
										{
											$SIPNATPrefixCollection = @() #null the collection for each table
											$SIPNATPrefixHeading = Fix-NullDescription -TableDescription $SIPNATPrefixTable.IE.Description -TableValue $SIPNATPrefixTable.value -TablePrefix 'SIP NAT Prefix Table ID: '
											$SIPNATPrefixEntries = $SIPNATPrefixTable.GetElementsByTagName('ID')
											ForEach ($SIPNATPrefixEntry in $SIPNATPrefixEntries)
											{
												if ($SIPNATPrefixEntry.IE.classname -eq $null) { continue } # Empty / deleted entry
												$SIPNATPrefixObject = @($SIPNATPrefixEntry.Value, $SIPNATPrefixEntry.IE.Description, $SIPNATPrefixEntry.IE.IPAddress, $SIPNATPrefixEntry.IE.Netmask)
												$SIPNATPrefixCollection += , $SIPNATPrefixObject
											}
											$AllSIPNATPrefixData += ,('NAT Qualified Prefix Tables', $SIPNATPrefixHeading, $SIPNATPrefixColumnTitles, $SIPNATPrefixCollection)
										}
									}
								}
							}

							# ---- SIP SIPSPRConditionRules (these are 'Message Manipulation / Condition Rule Tables' from the main menu) ---------------
							if ($SIPGroup.name -eq 'SIPSPRConditionRules')
							{
								$SIPCondRuleTables = $SIPGroup.GetElementsByTagName('ID')
								$SIPCondRuleColumnTitles = @('Entry', 'Description', 'Match Type', 'Operation', 'Match Value Type', 'Match Value')
								$SIPCondRuleCollection = @()
								$SIPCondRuleData = @()
								if ($SIPCondRuleTables.Count -ne 0)
								{
									ForEach ($SIPCondRuleTable in $SIPCondRuleTables)
									{
										if ($SIPCondRuleTable.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($SIPCondRuleTable.IE.classname -eq 'SPR_CONDITION_RULE_IE')
										{
											if ($SIPCondRuleTable.IE.classname -eq $null) { continue } # Empty / deleted entry
											$SIPCondRuleObject = @() #null the collection for each table
											$SIPCondRuleHeading = Fix-NullDescription -TableDescription $SIPCondRuleTable.IE.Description -TableValue $SIPCondRuleTable.value -TablePrefix 'SPR Condition Rule '
											$SIPCondRuleLookup.Add($SIPCondRuleTable.value, $SIPCondRuleHeading) # Referenced by SIP Msg Rule Tables
											if ($SIPCondRuleTable.IE.ConditionMatchType -eq '0')
											{
												$SIPCondRuleOperation = '<n/a>'
											}
											else
											{
												$SIPCondRuleOperation = $SIPCondRuleMatchTypeLookup.Get_Item($SIPCondRuleTable.IE.Operand2Type)
											}
											$SIPCondRuleObject = @($SIPCondRuleTable.Value, $SIPCondRuleTable.IE.Description, $SIPCondRuleTable.IE.Operand1, $SIPCondRuleOperationLookup.Get_Item($SIPCondRuleTable.IE.ConditionMatchType), $SIPCondRuleOperation, $SIPCondRuleTable.IE.Operand2)
											$SIPCondRuleCollection += , $SIPCondRuleObject
										}
									}
									$SIPCondRuleData += ,('Message Manipulation', 'Condition Rules', $SIPCondRuleColumnTitles, $SIPCondRuleCollection)
								}
							}

							# ---- SIP SIPSPRMessageRules (these are 'Message Manipulation / Message rule Tables' from the main menu) ---------------
							if ($SIPGroup.name -eq 'SIPSPRMessageRules')
							{
								$AllSipMsgRuleData = @()
								$MessageRuleTables = $SIPGroup.GetElementsByTagName('ID')
								$MessageRuleTablesColumnTitles = @('Entry', 'Description', 'Result Type', 'Message Type')
								$MessageRuleTableCollection = @()
								if ($MessageRuleTables.Count -ne 0)
								{
									# The handling for Message Rule Tables is unique and warrants a brief explanation.
									# For these guys we need to run through this section of the XML backup *twice*. The first time we generate a
									# table for what the user sees when they first select the heading 'Message Rule Tables'.
									# Then we run through a second time to populate each of the individual tables, as they have new/different information
									# (and pull extra config from the ElementDescriptors that were read from the XML in the initial parse).
									ForEach ($MessageRuleTable in $MessageRuleTables)
									{
										if ($MessageRuleTable.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($MessageRuleTable.IE.classname -eq 'SPR_MESSAGE_TABLE_IE')
										{
											$MessageRuleTableHeading = Fix-NullDescription -TableDescription $MessageRuleTable.IE.Description -TableValue $MessageRuleTable.value -TablePrefix 'Message Rule Table #'
											if ($MessageRuleTable.IE.classname -eq $null) { continue } # Empty / deleted entry
											switch ($MessageRuleTable.IE.MessageType)
											{
												'0'
												{
													$MessageRuleMessageType = 'All Messages'
												}
												'1'
												{
													$MessageRuleMessageType = 'All Requests'
												}
												'2'
												{
													$MessageRuleMessageType = 'All Responses'
												}
												'3'
												{
													$MessageRuleMessageType = "Selected Messages:`n"
													$SelectedMessagesList = ($MessageRuleTable.IE.MessageTypeList).Split(',')
													foreach ($SelectedMessage in $SelectedMessagesList)
													{
														if ($SelectedMessage -match '\d\d\d')
														{
															$MessageRuleMessageType += $SIPDescriptionLookup.Get_Item($SelectedMessage) + "`n"
														}
														else
														{
															$MessageRuleMessageType += $SelectedMessage + "`n"
														}
													}
													$MessageRuleMessageType = Strip-TrailingCR -DelimitedString $MessageRuleMessageType
												}
											}
											$MessageRuleTableObject = @($MessageRuleTable.Value, $MessageRuleTableHeading, $MatchTypeLookup.Get_Item($MessageRuleTable.IE.ResultType), $MessageRuleMessageType)
											$MessageRuleTableCollection += , $MessageRuleTableObject
										}
									}
									$AllSipMsgRuleData += ,('Message Manipulation', 'Message Rule Tables - Overview', $MessageRuleTablesColumnTitles, $MessageRuleTableCollection)
									# ---------------------------------------------------------------------------------------
									# This is the second pass through the data. Now we're building a separate table for each Message Rule Table.
									# ---------------------------------------------------------------------------------------
									$MessageRulesColumnTitles = @('Entry', 'Enabled', 'Description', 'Rule Type', 'Result Type', 'Condition Expression', 'Actions')
									ForEach ($MessageRuleTable in $MessageRuleTables)
									{
										$MessageRuleTableObject = @()
										$MessageRuleTableCollection = @()

										if ($MessageRuleTable.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($MessageRuleTable.IE.classname -eq 'SPR_MESSAGE_TABLE_IE')
										{
											$MessageRuleTableHeading = Fix-NullDescription -TableDescription $MessageRuleTable.IE.Description -TableValue $MessageRuleTable.value -TablePrefix 'Message Rule Table #'
											if ($MessageRuleTable.IE.classname -eq $null) { continue } # Empty / deleted entry
											$MessageRules = $MessageRuleTable.GetElementsByTagName('ID')
											foreach ($MessageRule in $Messagerules)
											{
												#This is the individual message rule, inside the table
												if ($MessageRule.IE.classname -eq $null) { continue } # Empty / deleted entry

												#Re-init/ flush these all for each loop
												$ElementDescListArray = @()
												$ElementUriPListArray = @()
												$ElementUriUListArray = @()
												$ElementHeaderListArray = @()
												$ElementHeaderParamListArray = @()

												$MessageRuleHeading = Fix-NullDescription -TableDescription $MessageRule.IE.Description -TableValue $MessageRule.value -TablePrefix 'Message Rule #'
												if ($MessageRule.IE.ConditionExpression -eq '')
												{
													$MessageRuleCondExp = 'Always Match'
												}
												else
												{
													$MessageRuleCondExp = ("Match {0}`n`n(Numbers in braces are the Condition Rule Table numbers)" -f $MessageRule.IE.ConditionExpression)
												}
												$MessageRuleAction = ''
												switch ($MessageRule.IE.classname)
												{
													'SPR_REQUEST_LINE_RULE_IE'
													{
														$MessageRuleRuleType = 'Request Line'
														$ElementDescListArray = ($MessageRule.IE.ElementDescriptorList).Split(',',[System.StringSplitOptions]::RemoveEmptyEntries)
														$ElementUriPListArray = ($MessageRule.IE.URIParameterElementDescriptorList).Split(',',[System.StringSplitOptions]::RemoveEmptyEntries)
														$ElementUriUListArray = ($MessageRule.IE.URIUserParameterElementDescriptorList).Split(',',[System.StringSplitOptions]::RemoveEmptyEntries)
														$MessageRuleAction += Sequence-SmmElements $ElementDescListArray $ElementUriUListArray $ElementUriPListArray $ElementHeaderListArray $ElementHeaderParamListArray 'Request Line Value'
													}
													'SPR_HEADER_RULE_IE'
													{
														$MessageRuleRuleType = 'Header Rule'
														$MessageRuleAction += 'Header Action: ' + $SipElementDescActionLookup.Get_Item($MessageRule.IE.HeaderAction) + "`n"
														$MessageRuleAction += 'Header Name: ' + (Get-Culture).TextInfo.ToTitleCase($MessageRule.IE.HeaderName) + "`n"
														switch ($MessageRule.IE.HeaderName)
														{
															{($_ -eq 'Contact') -or ($_ -eq 'Route') -or ($_ -eq 'Record-Route') -or ($_ -eq 'History-Info') -or ($_ -eq 'P-Asserted-Identity')}
															{
																if (($MessageRule.IE.HeaderOrdinal -ne '') -and ($MessageRule.IE.HeaderAction -eq 2))
																{
																	#Ordinal only shows if Header Action = Modify and only for the above Headers.
																	$MessageRuleAction += 'Header Ordinal Number: ' + $SIPMsgRuleHdrOrdnlLookup.Get_Item($MessageRule.IE.HeaderOrdinal) + "`n"
																}
															}
														}
														$MessageRuleAction += "`n"
														$ElementHeaderListArray = ($MessageRule.IE.HeaderElementDescriptorList).Split(',',[System.StringSplitOptions]::RemoveEmptyEntries)
														$ElementHeaderParamListArray = ($MessageRule.IE.HeaderParameterElementDescriptorList).Split(',',[System.StringSplitOptions]::RemoveEmptyEntries)
														$ElementUriPListArray = ($MessageRule.IE.URIParameterElementDescriptorList).Split(',',[System.StringSplitOptions]::RemoveEmptyEntries)
														$ElementUriUListArray = ($MessageRule.IE.URIUserParameterElementDescriptorList).Split(',',[System.StringSplitOptions]::RemoveEmptyEntries)
														$MessageRuleAction += Sequence-SmmElements $ElementDescListArray $ElementUriUListArray $ElementUriPListArray $ElementHeaderListArray $ElementHeaderParamListArray 'Header Value'
													}
													'SPM_RAW_MESSAGE_RULE_IE'
													{
														$MessageRuleRuleType = 'Raw Message Rule'
														$MessageRuleAction += 'Match Regex: ' + $MessageRule.IE.MatchRegex + "`nReplace Regex: " + $MessageRule.IE.ReplaceRegex
													}
													'SPR_PAYLOAD_RULE_IE'
													{
														$MessageRuleRuleType = 'Status Line Rule'
														$ElementDescListArray = ($MessageRule.IE.ElementDescriptorList).Split(',',[System.StringSplitOptions]::RemoveEmptyEntries)
														#These rules don't need complicated re-ordering:
														foreach ($ElementListItem in $ElementDescListArray)
														{
															$MessageRuleAction += $SIPMessageRuleElementLookup.Get_Item($ElementListItem) + "`n"
														}
														$MessageRuleAction = [regex]::replace($MessageRuleAction, '<Default>' , 'Status Line Value')
													}
													default { continue }
												}
												$MessageRuleAction = Strip-TrailingCR -DelimitedString $MessageRuleAction
												$MessageRuleTableObject = @($MessageRule.value, $EnabledLookup.Get_Item($MessageRule.IE.Enabled), $MessageRuleHeading, $MessageRuleRuleType, $MatchTypeLookup.Get_Item($MessageRule.IE.ResultType), $MessageRuleCondExp, $MessageRuleAction)
												$MessageRuleTableCollection += , $MessageRuleTableObject
											}
											#Now we need to rearrange the entries in the table into their correct sequence:
											#(Sequence is specified as an absolute, rather than encoded, as all previous sequences have been. e.g. it's now '3001,2,1')
											if (($MessageRuleTable.IE.Sequence -ne '') -and ($MessageRuleTable.IE.Sequence -ne $null)) #If it exists AND it's not blank
											{
												$MessageRuleCollectionTemp =@()
												$MessageRuleSequence = ($MessageRuleTable.IE.Sequence).Split(',')
												foreach ($reindex in $MessageRuleSequence)
												{
													foreach ($tableRow in $MessageRuleTableCollection)
													{
														if ($tableRow[0] -eq $reindex) # (0 is the 'Index' value)
														{
															$MessageRuleCollectionTemp += ,$TableRow
														}
													}
												}
												$MessageRuleTableCollection = $MessageRuleCollectionTemp
											}
											$AllSipMsgRuleData += ,('Message Manipulation', $MessageRuleTableHeading, $MessageRulesColumnTitles, $MessageRuleTableCollection)
										}
									}
								}
							}

							#---------- Node-Level SIP Settings ------------------------
							if ($SIPGroup.name -eq 'SIPCfgCACProfile')
							{
								$NodeLevelSIPObjects = $node.GetElementsByTagName('ID')
								$NodeLevelSIPSettings = @()
								if ($NodeLevelSIPObjects.Count -ne 0)
								{
									ForEach ($NodeLevelSIPObject in $NodeLevelSIPObjects)
									{
										if ($NodeLevelSIPObject.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($NodeLevelSIPObject.IE.classname -eq 'SIP_CFG_CAC_PROFILE_IE')
										{
											$NodeLevelSIPSettings += ,('SPAN-L', 'Node-Level SIP Settings', '', '')
											$NodeLevelSIPSettings += ,('SPAN-L', 'Call Admission Control', '', '')
											$NodeLevelSIPSettings += ,('CAC Admin State', (Test-ForNull -LookupTable $EnabledLookup -value $NodeLevelSIPObject.IE.EnableCAC), '', '')
											$NodeLevelSIPSettings += ,('SPAN-L', 'Source IP Untrusted Alarm', '', '')
											$NodeLevelSIPSettings += ,('Alarm for Untrusted INVITE', (Test-ForNull -LookupTable $EnabledLookup -value $NodeLevelSIPObject.IE.SourceIPUntrustedAlarm), '', '')
											$NodeLevelSIPSettings += ,('SPAN-L', 'Skype/Lync Presence', '', '')
											$NodeLevelSIPSettings += ,('Presence Destination', (Test-ForNull -LookupTable $SIPServerTablesLookup -value $NodeLevelSIPObject.IE.SkypePresenceClusterId), '', '')
											if (($NodeLevelSIPObject.IE.SkypePresenceClusterId -eq $null) -or ($NodeLevelSIPObject.IE.SkypePresenceClusterId -eq 0))
											{
												# Don't display Presence Server status if it's $null or 0
											}
											else
											{
												$NodeLevelSIPSettings += ,('Presence Service Status', '<Not Available>', '', '')
												$NodeLevelSIPSettings += ,('Publish Session Count', '<Not Available>', '', '')
											}
											if ($SweLite)
											{
												#The 'Skype/Lync Edge Server' section is not displaying in the SweLite as at v7.0.0 - but present in the config
											}
											else
											{
												$NodeLevelSIPSettings += ,('SPAN-L', 'Skype/Lync Edge Server', '', '')
												$NodeLevelSIPSettings += ,('Edge Server Destination', (Test-ForNull -LookupTable $SIPServerTablesLookup -value $NodeLevelSIPObject.IE.SkypeEdgeServerClusterId), '', '')
												if ($NodeLevelSIPObject.IE.TokenDuration -eq $null)
												{
													$NodeLevelSIPSettings += ,('Edge Server Token Duration', '<n/a this rls>', '', '')
												}
												else
												{
													$NodeLevelSIPSettings += ,('Edge Server Token Duration', ($NodeLevelSIPObject.IE.TokenDuration + ' [60..480] mins'), '', '')
												}
											}
											$NodeLevelSIPSettings += ,('SPAN-L', 'UserInfo Handling', '', '')
											$NodeLevelSIPSettings += ,('UserInfo Decode', (Test-ForNull -LookupTable $SIPUserInfoDecodeLookup -value $NodeLevelSIPObject.IE.UserInfoDecodePref), '', '')
											$NodeLevelSIPData += ,('Node-Level SIP Settings', '', '', $NodeLevelSIPSettings)
										}
									}
								}
							}
						}
					}

					#---------- Radius ------------------------
					'RAD'
					{
						$RadiusClassCollection  = @()
						$RadiusServerCollection = @()
						# ---- Configuration ----------
						if ($node.IE.Classname -eq 'RAD_RADIUS_CONFIG')
						{
							if ($node.IE.Enabled -eq '1')
							{
								$RadConfigTable = @()
								$RadConfigTable += ,('SPAN-L', 'RADIUS Options', '', '')
								$RadConfigTable += ,('Authentication', $EnabledLookup.Get_Item($node.IE.Authentication), '', '')
								$RadConfigTable += ,('Accounting', $EnabledLookup.Get_Item($node.IE.Accounting), '', '')
								if ($node.IE.Accounting -eq 1)
								{
									$RadConfigTable += ,('Accounting Mode', $RadiusAccountingModeLookup.Get_Item($node.IE.Accounting), '', '')
								}
								if ($node.IE.Authentication -eq 1)
								{
									$RadConfigTable += ,('SPAN-L', 'RADIUS Authentication Servers', '', '')
									$RadConfigTable += ,('Primary Server', $RadiusServerLookup.Get_Item($node.IE.PrimaryRADServer) , '', '')
									$RadConfigTable += ,('Secondary Server', $RadiusServerLookup.Get_Item($node.IE.SecondaryRADServer) , '', '')
								}
								if ($node.IE.Accounting -eq 1)
								{
									$RadConfigTable += ,('SPAN-L', 'RADIUS Accounting Servers', '', '')
									$RadConfigTable += ,('Primary Server', $RadiusServerLookup.Get_Item($node.IE.AccountingServer1) , '', '')
									$RadConfigTable += ,('Secondary Server', $RadiusServerLookup.Get_Item($node.IE.AccountingServer2) , '', '')
								}
								$ADData += , ('RADIUS', 'Configuration', '', $RadConfigTable)
							}
						}
						$RadiusUserGroups = $node.GetElementsByTagName('Token')
						ForEach ($RadiusUserGroup in $RadiusUserGroups)
						{
							# ---- Radius Servers ----------
							if ($RadiusUserGroup.name -eq 'Server')
							{
								$RadiusServers = $RadiusUserGroup.GetElementsByTagName('ID')
								$RadiusServerColumnTitles = @('Description', 'RADIUS Server', 'Port', 'Send Connectivity Packet Type')
								if ($RadiusServers.Count -ne 0)
								{
									ForEach ($RadiusServer in $RadiusServers)
									{
										if ($RadiusServer.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($RadiusServer.IE.classname -eq 'RAD_RADIUS_SERVERS')
										{
											if ($RadiusServer.IE.SendConnectivityPacket -eq '0')
											{
												$RadiusServerObject = @($RadiusServer.IE.Description, $RadiusServer.IE.RadiusServer, $RadiusServer.IE.Port, 'Status-Server (RFC 5997)')
											}
											else
											{
												$RadiusServerObject = @($RadiusServer.IE.Description, $RadiusServer.IE.RadiusServer, $RadiusServer.IE.Port, 'Request')
											}
											$RadiusServerCollection += , $RadiusServerObject
										}
									}
									$ADData += , ('RADIUS', 'RADIUS Servers', $RadiusServerColumnTitles, $RadiusServerCollection)
								}
							}
							# ---- Radius User Groups ----------
							if ($RadiusUserGroup.name -eq 'Mapping')
							{
								if ($RadiusUserGroup.IE.classname -eq 'UAS_USER_CLASS')
								{
									$RadiusClassObject = @('[Missing User Class Access Level]', $AccountAccessLevelLookup.Get_Item($RadiusUserGroup.IE.AuthLevel))
									$RadiusClassCollection += , $RadiusClassObject
								}
								$RadiusClasses = $RadiusUserGroup.GetElementsByTagName('ID')
								$RadiusClassColumnTitles = @('Class Name', 'Access Level', 'Primary Key')
								if ($RadiusClasses.Count -ne 0)
								{
									ForEach ($RadiusClass in $RadiusClasses)
									{
										if ($RadiusClass.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($RadiusClass.IE.classname -eq 'UAS_USER_CLASS')
										{
											#Special handling was needed here as there's an 'element' and an 'attribute' at the same level, both cleverly called ClassName
											$ClassName = $RadiusClass.selectsinglenode('IE/ClassName')
											$ClassNameText = $ClassName.'#text'
											$RadiusClassObject = @($ClassNameText, $AccountAccessLevelLookup.Get_Item($RadiusClass.IE.AuthLevel), $RadiusClass.value)
											$RadiusClassCollection += , $RadiusClassObject
										}
									}
									#The $RadiusClassCollection is added to $SecurityData just prior to printing to ensure the correct ordering in the DOCX
								}
							}
						}
					}

					#---------- CAS ------------------------
					'CAS'
					{
						$CASProfileGroups = $node.GetElementsByTagName('Token')
						$CASData = @()

						ForEach ($CASProfileGroup in $CASProfileGroups)
						{
							# ---- CASLoopStartFxs Profiles ----------
							if ($CASProfileGroup.name -eq 'CASLoopStartFxs')
							{
								$CasProfilesFxsProfiles = $CASProfileGroup.GetElementsByTagName('ID')
								If ($CasProfilesFxsProfiles.Count -ne 0)
								{
									ForEach ($CasProfilesFxsProfile in $CasProfilesFxsProfiles)
									{
										if ($CasProfilesFxsProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($CasProfilesFxsProfile.IE.classname -eq 'CAS_LOOPSTART_FXS_CFG_IE')
										{
											$CasProfilesFxsTable = @() #null the collection for each table
											$CasProfilesFxsHeading = Fix-NullDescription -TableDescription $CasProfilesFxsProfile.IE.Description -TableValue $CasProfilesFxsProfile.value -TablePrefix 'CAS Profile #'
											$CasProfilesFxsHeading = '(FXS) ' + $CasProfilesFxsHeading
											#Build the table:
											$CasProfilesFxsTable += ,('SPAN-L', 'CAS Signaling Profile Table', '' , '')
											$CasProfilesFxsTable += ,('Description', $CasProfilesFxsHeading, '' , '')
											$CasProfilesFxsTable += ,('SPAN-L', 'Loop Start FXS Properties', '' , '')
											switch ($CasProfilesFxsProfile.IE.LoopStartType)
											{
												'0'
												{
													$CasProfilesFxsTable += ,('Loop Start Type', 'Basic', '' , '')

												}
												'1'
												{
													$CasProfilesFxsTable += ,('Loop Start Type', 'Reverse Battery', '' , '')
													if ($CasProfilesRTWOProfile.IE.RevBattIncomingImmediate -ne $null)
													{
														$CasProfilesFxsTable += ,('Incoming Immediate', $EnabledLookup.Get_Item($CasProfilesFxsProfile.IE.RevBattIncomingImmediate), '' , '')
													}
													else
													{
														$CasProfilesFxsTable += ,('Incoming Immediate', '<n/a this rls>', '', '')
													}
												}
												'2'
												{
													$CasProfilesFxsTable += ,('Loop Start Type', 'Forward Disconnect', '' , '')
													$CasProfilesFxsTable += ,('Forward Disconnect Duration', ($CasProfilesFxsProfile.IE.ForwardDisconnectDuration + ' ms'), '' , '')
												}
											}
											$CasProfilesFxsTable += ,('Disconnect Tone Generation', $EnabledLookup.Get_Item($CasProfilesFxsProfile.IE.DisconnectToneGeneration), '' , '')
											if (($CasProfilesFxsProfile.IE.FlashhookSignalDetection) -eq 1)
											{
												$CasProfilesFxsTable += ,('Flashhook Signal Detection', 'Enabled', '' , '')
												$CasProfilesFxsTable += ,('Maximum Flashhook Duration', ($CasProfilesFxsProfile.IE.MaximumFlashhookDuration + ' ms'), '' , '')
												$CasProfilesFxsTable += ,('Minimum Flashhook Duration', ($CasProfilesFxsProfile.IE.MinimumFlashhookDuration + ' ms'), '' , '')
											}
											else
											{
												$CasProfilesFxsTable += ,('Flashhook Signal Detection', 'Disabled', '' , '')
											}
											$CasProfilesFxsTable += ,('Inter-Digit Timeout', ($CasProfilesFxsProfile.IE.InterDigitTimeout + ' ms'), '' , '')
											$CasProfilesFxsTable += ,('SPAN-L', 'Ringing Cadence', '' , '')

											$CasProfilesFxsTable += ,('Cadence On', ($CasProfilesFxsProfile.IE.CadenceOnTime + ' ms'), '' , '')
											$CasProfilesFxsTable += ,('Cadence Off', ($CasProfilesFxsProfile.IE.CadenceOffTime + ' ms'), '' , '')

											if ($CasProfilesFxsProfile.IE.DoubleCadence -eq 1)
											{
												$CasProfilesFxsTable += ,('Double Cadence', 'Yes', '' , '')
												$CasProfilesFxsTable += ,('Cadence On', ($CasProfilesFxsProfile.IE.CadenceOnTime2 + ' ms'), '' , '')
												$CasProfilesFxsTable += ,('Cadence On', ($CasProfilesFxsProfile.IE.CadenceOffTime2 + ' ms'), '' , '')
											}
											else
											{
												$CasProfilesFxsTable += ,('Double Cadence', 'No', '' , '')
											}
											$CASData += ,('CAS Signaling Profiles', $CasProfilesFxsHeading, '', $CasProfilesFxsTable)
										}
									}
								}
							}

							# ---- CASLoopStartFXO Profiles ----------
							if ($CASProfileGroup.name -eq 'CASLoopStartFXO')
							{
								$CasProfilesFxoProfiles = $CASProfileGroup.GetElementsByTagName('ID')
								If ($CasProfilesFxoProfiles.Count -ne 0)
								{
									ForEach ($CasProfilesFxoProfile in $CasProfilesFxoProfiles)
									{
										if ($CasProfilesFxoProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($CasProfilesFxoProfile.IE.classname -eq 'CAS_LOOPSTART_FXO_CFG_IE')
										{
											$CasProfilesFxoTable = @() #null the collection for each table
											$CasProfilesFxoHeading = Fix-NullDescription -TableDescription $CasProfilesFxoProfile.IE.Description -TableValue $CasProfilesFxoProfile.value -TablePrefix 'CAS Profile #'
											$CasProfilesFxoHeading = '(FXO) ' + $CasProfilesFxoHeading
											#Build the table:
											$CasProfilesFxoTable += ,('SPAN-L', 'CAS Signaling Profile Table', '' , '')
											$CasProfilesFxoTable += ,('Description', $CasProfilesFxoHeading, '' , '')
											$CasProfilesFxoTable += ,('SPAN-L', 'Loop Start FXO Properties', '' , '')
											switch ($CasProfilesFxoProfile.IE.LoopStartType)
											{
												'0'
												{
													$CasProfilesFxoTable += ,('Loop Start Type', 'Basic', '' , '')
													$CasProfilesFxoTable += ,('DTMF On Time', $CasProfilesFxoProfile.IE.DTMFOnTime, '' , '')
													$CasProfilesFxoTable += ,('DTMF Off Time', $CasProfilesFxoProfile.IE.DTMFOffTime, '' , '')

												}
												'1'
												{
													$CasProfilesFxoTable += ,('Loop Start Type', 'Reverse Battery', '' , '')
													$CasProfilesFxoTable += ,('DTMF On Time', $CasProfilesFxoProfile.IE.DTMFOnTime, '' , '')
													$CasProfilesFxoTable += ,('DTMF Off Time', $CasProfilesFxoProfile.IE.DTMFOffTime, '' , '')
												}
												'2'
												{
													$CasProfilesFxoTable += ,('Loop Start Type', 'Forward Disconnect', '' , '')
													$CasProfilesFxoTable += ,('DTMF On Time', $CasProfilesFxoProfile.IE.DTMFOnTime, '' , '')
													$CasProfilesFxoTable += ,('DTMF Off Time', $CasProfilesFxoProfile.IE.DTMFOffTime, '' , '')
													$CasProfilesFxoTable += ,('Min. Forward Disconnect Duration', ($CasProfilesFxoProfile.IE.MinimumForwardDisconnectDuration + ' ms'), '' , '')
												}
											}
											if ($CasProfilesFxoProfile.IE.SilenceBasedDisconnect -eq '0')
											{
												$CasProfilesFxoTable += ,('Silence Based Disconnect', 'Disabled', '' , '')
											}
											else
											{
												$CasProfilesFxoTable += ,('Silence Based Disconnect', 'Enabled', '' , '')
												$CasProfilesFxoTable += ,('Silence Level', ($CasProfilesFxoProfile.IE.SilenceDetectionLevel + ' dBm'), '' , '')
												$CasProfilesFxoTable += ,('Silence Based Disconnect Timer', ($CasProfilesFxoProfile.IE.SilenceBasedDisconnect + ' sec.'), '' , '')
											}
											if ($CasProfilesFxoProfile.IE.LoopStartType -eq '0')
											{
												if ($CasProfilesFxoProfile.IE.DisconnectSupervision -eq '0')
												{
													$CasProfilesFxoTable += ,('Tone Based Disconnect', 'Disabled', '' , '')
												}
												else
												{
													$CasProfilesFxoTable += ,('Tone Based Disconnect', 'Enabled', '' , '')
													$CasProfilesFxoTable += ,('Disconnect Tone Cadences Per Cycle', $CasProfilesFxoProfile.IE.DisconnectSupervisionOptions , '' , '')
												}
											}
											if ($CasProfilesFxoProfile.IE.DialToneDetect -eq '0')
											{
												$CasProfilesFxoTable += ,('Dial Tone Detect', 'Disabled', '' , '')
												$CasProfilesFxoTable += ,('Dial Delay', ($CasProfilesFxoProfile.IE.DialDelay + ' ms'), '' , '')
											}
											else
											{
												$CasProfilesFxoTable += ,('Dial Tone Detect', 'Enabled', '' , '')
											}
											$CasProfilesFxoTable += ,('Provide Tone With Answer', $EnabledLookup.Get_Item($CasProfilesFxoProfile.IE.ProvideToneWithAnswer), '' , '')
											$CASData += ,('CAS Signaling Profiles', $CasProfilesFxoHeading, '', $CasProfilesFxoTable)
										}
									}
								}
							}

							# ---- CAS E&M Profiles ----------
							if ($CASProfileGroup.name -eq 'CASEnM')
							{
								$CasProfilesENMProfiles = $CASProfileGroup.GetElementsByTagName('ID')
								If ($CasProfilesENMProfiles.Count -ne 0)
								{
									ForEach ($CasProfilesENMProfile in $CasProfilesENMProfiles)
									{
										if ($CasProfilesENMProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($CasProfilesENMProfile.IE.classname -eq 'CAS_ENM_CFG_IE')
										{
											$CasProfilesENMTable = @() #null the collection for each table
											$CasProfilesENMHeading = Fix-NullDescription -TableDescription $CasProfilesENMProfile.IE.Description -TableValue $CasProfilesENMProfile.value -TablePrefix 'CAS Profile #'
											$CasProfilesENMHeading = '(E&M) ' + $CasProfilesENMHeading
											#Build the table:
											$CasProfilesENMTable += ,('SPAN-L', 'CAS Signaling Profile Table', '' , '')
											$CasProfilesENMTable += ,('Description', $CasProfilesENMHeading, '' , '')
											$CasProfilesENMTable += ,('Orientation', $CasOrientationLookup.Get_Item($CasProfilesENMProfile.IE.Orientation), '' , '')
											$CasProfilesENMTable += ,('SPAN-L', 'E&M Signaling', '' , '')

											$CasProfilesENMTable += ,('Incoming Start Dial', $CasStartDialLookup.Get_Item($CasProfilesENMProfile.IE.IncomingStartDial), '' , '')
											$CasProfilesENMTable += ,('Outgoing Start Dial', $CasStartDialLookup.Get_Item($CasProfilesENMProfile.IE.OutgoingStartDial), '' , '')
											$CasProfilesENMTable += ,('DTMF On', ($CasProfilesENMProfile.IE.DTMFOnTime + ' ms'), '' , '')
											$CasProfilesENMTable += ,('DTMF Off', ($CasProfilesENMProfile.IE.DTMFOffTime + ' ms'), '' , '')
											$CasProfilesENMTable += ,('Inter-Digit Timeout', ($CasProfilesENMProfile.IE.InterDigitTimeout + ' ms'), '' , '')

											$CASData += ,('CAS Signaling Profiles', $CasProfilesENMHeading, '', $CasProfilesENMTable)
										}
									}
								}
							}
							# ---- CAS R2MFC Profiles ----------
							if ($CASProfileGroup.name -eq 'CASR2')
							{
								$CasProfilesRTWOProfiles = $CASProfileGroup.GetElementsByTagName('ID')
								#$CASLoopRTWOProfilesColumnTitles = @('Entry', 'Description', 'Orientation', 'Incoming Signal Type', 'Outgoing Signal Type' )
								If ($CasProfilesRTWOProfiles.Count -ne 0)
								{
									ForEach ($CasProfilesRTWOProfile in $CasProfilesRTWOProfiles)
									{
										if ($CasProfilesRTWOProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($CasProfilesRTWOProfile.IE.classname -eq 'CAS_RTWO_CFG_IE')
										{
											$CasProfilesRTWOTable = @() #null the collection for each table
											$CasProfilesRTWOHeading = Fix-NullDescription -TableDescription $CasProfilesRTWOProfile.IE.Description -TableValue $CasProfilesRTWOProfile.value -TablePrefix 'CAS Profile #'
											$CasProfilesRTWOHeading = '(R2) ' + $CasProfilesRTWOHeading

											#--- NEW:
											#Build the table:
											$CasProfilesRTWOTable += ,('SPAN-L', 'CAS Signaling Profile Table', '' , '')
											$CasProfilesRTWOTable += ,('Description', $CasProfilesRTWOHeading, '' , '')
											$CasProfilesRTWOTable += ,('Orientation', $CasOrientationLookup.Get_Item($CasProfilesRTWOProfile.IE.Orientation), '' , '')
											$CasProfilesRTWOTable += ,('Incoming Signal Type', $CasR2SignalingTypeLookup.Get_Item($CasProfilesRTWOProfile.IE.IncomingTone), '' , '')
											$CasProfilesRTWOTable += ,('Outgoing Signal Type', $CasR2SignalingTypeLookup.Get_Item($CasProfilesRTWOProfile.IE.OutgoingTone), '' , '')
											$CasProfilesRTWOTable += ,('SPAN-L', 'Line Signaling Options', '' , '')
											$CasProfilesRTWOTable += ,('CD Bits', $CasR2CDBitsLookup.Get_Item($CasProfilesRTWOProfile.IE.CDBits), '' , '')
											if ($CasProfilesRTWOProfile.IE.InvertABCDBits -eq '1')
											{
												$CasProfilesRTWOTable += ,('Inverted ABCD Bits', $CasR2InvertedABCDBitsLookup.Get_Item($CasProfilesRTWOProfile.IE.InvertedABCDBits), '' , '')
											}
											else
											{
												$CasProfilesRTWOTable += ,('Inverted ABCD Bits', 'None', '' , '')
											}
											$CasProfilesRTWOTable += ,('Meter Code', $EnabledLookup.Get_Item($CasProfilesRTWOProfile.IE.MeterCode), '' , '')
											$CasProfilesRTWOTable += ,('Release Guard Time', ($CasProfilesRTWOProfile.IE.ReleaseGuardTime + ' ms'), '' , '')
											$CasProfilesRTWOTable += ,('Seizure Acknowledgement Timer', ($CasProfilesRTWOProfile.IE.SeizureAcknowledgmentTime + ' ms'), '' , '')
											$CasProfilesRTWOTable += ,('Inter-Digit Timeout', ($CasProfilesRTWOProfile.IE.InterDigitTimeout + ' ms'), '' , '')
											if (($CasProfilesRTWOProfile.IE.IncomingTone -eq '1') -or ($CasProfilesRTWOProfile.IE.OutgoingTone -eq '1'))
											{
												#Things get a LOT more complicated if either i/c or o/g is MFC:
												$CasProfilesRTWOTable += ,('SPAN-L', 'Register Signaling Options', '' , '')
												if ($CasProfilesRTWOProfile.IE.RequestANI -eq '1')
												{
													$CasProfilesRTWOTable += ,('Request Calling Party Number', 'Enabled', '', '')
													$CasProfilesRTWOTable += ,('Digits Received Before Request', ($CasR2DigitsRxdLookup.Get_Item($CasProfilesRTWOProfile.IE.DNISDigits2RequestANI)+ " (Called Party digits rec'd)"), '', '')
													if ($CasProfilesRTWOProfile.IE.VariableANILength -eq '0')
													{
														$CasProfilesRTWOTable += ,('Calling Party Number Length', 'Fixed', '', '')
														$CasProfilesRTWOTable += ,('Calling Party Digits', $CasProfilesRTWOProfile.IE.ANILength, '', '')
													}
													else
													{
														$CasProfilesRTWOTable += ,('Calling Party Number Length', 'Variable', '', '')
													}
												}
												else
												{
													$CasProfilesRTWOTable += ,('Request Calling Party Number', 'Disabled', '', '')
												}
												if ($CasProfilesRTWOProfile.IE.VariableDNISLength -eq '0')
												{
													$CasProfilesRTWOTable += ,('Called Party Number Length', 'Fixed (Incoming calls)', '', '')
													$CasProfilesRTWOTable += ,('Called Party Digits', $CasProfilesRTWOProfile.IE.DNISLength, '', '')
												}
												else
												{
													$CasProfilesRTWOTable += ,('Called Party Number Length', 'Variable (Incoming calls)', '', '')
												}
												if ($CasProfilesRTWOProfile.IE.SendEndOfDigit -ne $null)
												{
													$CasProfilesRTWOTable += ,('Send End Of Digit', ($EnabledLookup.Get_Item($CasProfilesRTWOProfile.IE.SendEndOfDigit) + ' (Outgoing calls)'), '', '')
												}
												else
												{
													$CasProfilesRTWOTable += ,('Send End Of Digit', '<n/a this rls>', '', '')
												}
												$CasProfilesRTWOTable += ,('SPAN-L', 'Group III/C Tones', '' , '')
												if ($CasProfilesRTWOProfile.IE.UseGroupCTones -ne $null)
												{
													$CasProfilesRTWOTable += ,('Use Group-C Tones', $EnabledLookup.Get_Item($CasProfilesRTWOProfile.IE.UseGroupCTones), '' , '')
												}
												else
												{
													$CasProfilesRTWOTable += ,('Use Group-C Tones', '<n/a this rls>', '', '')
												}
												$CasProfilesRTWOTable += ,('SPAN-L', 'Forward Register Signals', '' , '')
												$CasProfilesRTWOTable += ,("Calling Party's Category on A3", $CasR2CategoryLookup.Get_Item($CasProfilesRTWOProfile.IE.Group2Tone4A3Category), '' , '')
												$CasProfilesRTWOTable += ,("Calling Party's Category on Send CPC", ($CasR2CategoryLookup.Get_Item($CasProfilesRTWOProfile.IE.Group1Tone4A6Category) + ' (China Group KA)'), '' , '')
												if ($CasProfilesRTWOProfile.IE.Group1AniRestricted -ne $null)
												{
													$CasProfilesRTWOTable += ,('Group-I Tone', $CasR2Group1Lookup.Get_Item($CasProfilesRTWOProfile.IE.Group1AniRestricted), '' , '')
												}
												else
												{
													$CasProfilesRTWOTable += ,('Group-I Tone', '<n/a this rls>', '', '')
												}
												if ($CasProfilesRTWOProfile.IE.Group2AniRestricted -ne $null)
												{
													$CasProfilesRTWOTable += ,('Group-II Tone (CPC)', $CasR2CategoryLookup.Get_Item($CasProfilesRTWOProfile.IE.Group2AniRestricted), '' , '')
												}
												else
												{
													$CasProfilesRTWOTable += ,('Group-II Tone (CPC)', '<n/a this rls>', '', '')
												}
												$CasProfilesRTWOTable += ,('SPAN-L', 'Backward Register Signals', '' , '')
												$CasProfilesRTWOTable += ,('SPAN-L', 'Group A Signals', '', '' )
												$CasProfilesRTWOTable += ,("Send Calling Party's Number", $CasR2GroupASignalsLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupATone4RequestANIDigit), '', '')
												$CasProfilesRTWOTable += ,("Send Calling Party's Category", $CasR2GroupASignalsLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupATone4RequestCategoryDigit), '', '')
												if ($CasProfilesRTWOProfile.IE.GroupATone4AnswerDigit -eq '3')
												{
													$CasProfilesRTWOTable += ,('Address Complete (Changeover to Group B)', 'A-3 (Yes)', '', '')
												}
												else
												{
													$CasProfilesRTWOTable += ,('Address Complete (Changeover to Group B)', 'A-6 (No)', '', '')
												}
												$CasProfilesRTWOTable += ,('Send First Address Digit', $CasR2GroupASignalsLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupATone4Send1stDigit), '', '')
												$CasProfilesRTWOTable += ,('SPAN-L', 'Group B Signals', '', '' )
												$CasProfilesRTWOTable += ,('Subscriber Line Free Sent', $CasR2GroupBSignalsLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupBTone4IdleSent), '', '')
												$CasProfilesRTWOTable += ,('Subscriber Line Busy Sent', $CasR2GroupBSignalsLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupBTone4BusySent), '', '')
												if ($CasProfilesRTWOProfile.IE.GroupBTone4CongSent -ne $null)
												{
													$CasProfilesRTWOTable += ,('Congestion Sent', $CasR2GroupBSignalsLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupBTone4CongSent), '', '')
												}
												else
												{
													$CasProfilesRTWOTable += ,('Congestion Sent', '<n/a this rls>', '', '')
												}
												if ($CasProfilesRTWOProfile.IE.GroupBTone4UnallocNumSent -ne $null)
												{
													$CasProfilesRTWOTable += ,('Unallocated Number Sent', $CasR2GroupBSignalsLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupBTone4UnallocNumSent), '', '')
												}
												else
												{
													$CasProfilesRTWOTable += ,('Unallocated Number Sent', '<n/a this rls>', '', '')
												}
												$CasProfilesRTWOTable += ,('SPAN-L', 'Subscriber Line Free Received', '' , '')
												$CasProfilesRTWOTable += ,('B-1', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB1IdleReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-2', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB2IdleReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-3', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB3IdleReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-4', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB4IdleReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-5', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB5IdleReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-6', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB6IdleReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-7', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB7IdleReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-8', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB8IdleReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('SPAN-L', 'Subscriber Line Busy Received', '' , '')
												$CasProfilesRTWOTable += ,('B-1', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB1BusyReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-2', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB2BusyReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-3', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB3BusyReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-4', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB4BusyReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-5', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB5BusyReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-6', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB6BusyReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-7', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB7BusyReceivedBit), '', '')
												$CasProfilesRTWOTable += ,('B-8', $CasR2GroupBSignalsCheckLookup.Get_Item($CasProfilesRTWOProfile.IE.GroupB8BusyReceivedBit), '', '')
											}
											$CASData += ,('CAS Signaling Profiles', $CasProfilesRTWOHeading, '', $CasProfilesRTWOTable)
										}
									}
								}
							}
							# ---- CASSupplementary ----------
							if ($CASProfileGroup.name -eq 'CASSupplementary')
							{
								$CASSuppServicesProfiles = $CASProfileGroup.GetElementsByTagName('ID')
								$CASSuppServicesProfilesColumnTitles = @('Entry', 'Description', 'Hold', 'Transfer', 'Call Waiting', 'Conference')
								$CASSuppServicesProfileData = @()
								$CASSuppServicesProfileCollection = @()
								if ($CASSuppServicesProfiles.Count -ne 0)
								{
									ForEach ($CASSuppServicesProfile in $CASSuppServicesProfiles)
									{
										if ($CASSuppServicesProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($CASSuppServicesProfile.IE.classname -eq 'CAS_SUPPLEMENTARY_CFG_IE')
										{
											$CASSuppServicesHeading = Fix-NullDescription -TableDescription $CASSuppServicesProfile.IE.Description -TableValue $CASSuppServicesProfile.value -TablePrefix 'CAS Supplementary Service Profile #'
											$CASSuppServicesProfileObject = @($CASSuppServicesProfile.Value, $CASSuppServicesProfile.IE.Description, $EnabledLookup.Get_Item($CASSuppServicesProfile.IE.Hold),$EnabledLookup.Get_Item($CASSuppServicesProfile.IE.Transfer),$EnabledLookup.Get_Item($CASSuppServicesProfile.IE.CallWaiting),$EnabledLookup.Get_Item($CASSuppServicesProfile.IE.Conference))
											$CASSuppServicesProfileCollection += , $CASSuppServicesProfileObject
										}
									}
									$CASData += ,('CAS Supplementary Service Profiles', '', $CASSuppServicesProfilesColumnTitles, $CASSuppServicesProfileCollection)
								}
							}
						}
					}

					#---------- Emergency Services ------------------------
					'EmergencyServices'
					{
						$EmergSvcsGroups = $node.GetElementsByTagName('Token')
						ForEach ($EmergSvcsGroup in $EmergSvcsGroups)
						{
							# ---- Emergency Svcs Config ----------
							if ($EmergSvcsGroup.name -eq 'ESConfig')
							{
								if ($EmergSvcsGroup.IE.classname -eq 'CC_CFG_EMERGENCY_SERVICES_IE')
								{
									$EmergSvcsConfigTable = @() #null the collection for each table
									$EmergSvcsConfigTable += ,('SPAN-L', 'Emergency Services Configuration', '' , '')
									$EmergSvcsConfigTable += ,('SPAN-L', 'Set Call Status Duration', '' , '')
									$EmergSvcsConfigTable += ,('Call Status Duration', ($EmergSvcsGroup.IE.CallStatusDuration + ' hours'), '' , '')
									$EmergSvcsConfigTable += ,('SPAN-L', 'Call Preemption', '' , '')
									$EmergSvcsConfigTable += ,('Preempt Non-Emergency Calls', (Test-ForNull -LookupTable $EnabledLookup -value $EmergSvcsGroup.IE.PreemptNonEmergencyCalls), '' , '')
								}
								$EmergSvcsData += ,('Emergency Services Configuration', '', '', $EmergSvcsConfigTable)
							}

							# ---- Emergency Svcs Config ----------
							if ($EmergSvcsGroup.name -eq 'ESCallback')
							{
								$CallbackTableData = @()
								$CallbackTableProfiles = $EmergSvcsGroup.GetElementsByTagName('ID')
								if ($CallbackTableProfiles.Count -ne 0)
								{
									ForEach ($CallbackTableProfile in $CallbackTableProfiles)
									{
										if ($CallbackTableProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($CallbackTableProfile.IE.classname -eq 'CR_CALLBACK_NUMBER_TABLE_LIST_CONFIG')
										{
											$CallbackTableTable = @()
											$CallbackTableTable += ,('SPAN-L', 'Callback Numbers Table', '', '')
											$CallbackTableDescription = (Fix-NullDescription -TableDescription $CallbackTableProfile.IE.Description -TableValue $CallbackTableProfile.value -TablePrefix 'Callback Table #')
											$CallbackTableTable += ,('Description', $CallbackTableDescription, '' , '')
											$CallbackNumbersList = [regex]::replace($CallbackTableProfile.IE.CallbackNumbersList, ',', "`n")
											$CallbackTableTable += ,('Callback Numbers List', $CallbackNumbersList, '' , '')
											$CallbackTableTable += ,('PSAP Number', $CallbackTableProfile.IE.PSAPNumber, '' , '')
										}
										$EmergSvcsData += ,('Callback Number Pool', $CallbackTableDescription, '', $CallbackTableTable)
									}
								}
							}
						}
					}

					#---------- DHCP ------------------------
					'DHCP'
					{
						$DHCPGroups = $node.GetElementsByTagName('Token')
						$DHCPData = @()
						ForEach ($DHCPGroup in $DHCPGroups)
						{
							# ---- DHCP Pool ----------
							if ($DHCPGroup.name -eq 'DHCPPool')
							{
								$DHCPPools = $DHCPGroup.GetElementsByTagName('ID')
								$DHCPPoolsColumnTitles = @('Entry', 'Enabled', 'Name', 'Interface', 'Pool IP Start', 'Pool IP End', 'Lease Lifetime', 'Default Router', 'DNS Server')
								$DHCPPoolCollection = @()
								if ($DHCPPools.Count -ne 0)
								{
									ForEach ($DHCPPool in $DHCPPools)
									{
										if ($DHCPPool.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($DHCPPool.IE.classname -eq 'NETSERV_DHCP_OBJ_CFG_IE')
										{
											$DHCPPoolLookup.Add($DHCPPool.Value, $DHCPPool.IE.Name)	# Referenced by 'DHCPStaticClient' below
											if ($DHCPPool.IE.LeaseDuration -eq '2147483647')
											{
												$DHCPLeaseDuration = 'Never Expires'
											}
											else
											{
												$DHCPLeaseDuration = ([int32]$DHCPPool.IE.LeaseDuration / 60).ToString() + ' mins'
											}
											$DHCPPoolObject = @($DHCPPool.Value, $EnabledLookup.Get_Item($DHCPPool.IE.Enabled), $DHCPPool.IE.Name, $PortToIPAddressLookup.Get_Item($DHCPPool.IE.Interface), $DHCPPool.IE.PoolIPStart, $DHCPPool.IE.PoolIPEnd, $DHCPLeaseDuration, $DHCPPool.IE.DefaultRouterIP, $DHCPPool.IE.DnsServerIP)
											$DHCPPoolCollection += , $DHCPPoolObject
										}
									}
									$DHCPData += ,('DHCP', 'Pools', $DHCPPoolsColumnTitles, $DHCPPoolCollection) # On-screen, DHCP is part of IP Routing.
								}
							}
							# ---- DHCP Static Client ----------
							if ($DHCPGroup.name -eq 'DHCPStaticClient')
							{
								$DHCPClients = $DHCPGroup.GetElementsByTagName('ID')
								$DHCPClientsColumnTitles = @('Entry', 'MAC Address', 'Pool', 'IP Address', 'Lease Lifetime')
								$DHCPClientCollection = @()
								if ($DHCPClients.Count -ne 0)
								{
									ForEach ($DHCPClient in $DHCPClients)
									{
										if ($DHCPClient.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($DHCPClient.IE.classname -eq 'NETSERV_DHCP_OBJ_CFG_IE')
										{
											if ($DHCPClient.IE.LeaseDuration -eq '2147483647')
											{
												$DHCPClientLeaseDuration = 'Never Expires'
											}
											else
											{
												$DHCPClientLeaseDuration = ($DHCPClient.IE.LeaseDuration / 60).ToString() + ' mins'
											}
											$DHCPClientObject = @($DHCPClient.Value, $DHCPClient.IE.Name, $DHCPPoolLookup.Get_Item($DHCPClient.IE.ClientPoolID), $DHCPClient.IE.ClientIP, $DHCPClientLeaseDuration)
											$DHCPClientCollection += , $DHCPClientObject
										}
									}
									$DHCPData += ,('DHCP', 'Static Clients', $DHCPClientsColumnTitles, $DHCPClientCollection) # On-screen, DHCP is part of IP Routing.
								}
							}
						}
					}

					#---------- Global Security Options ------------------------
					'GlobalSecurityOptions'
					{
						$GlobalSecurityTable = @() # This will remain undeclared if the backup file is too old, and will be detected and discarded downstream
						if ($node.IE.Classname -eq 'UAS_GLOBAL_OPTS_CONFIG')
						{
							$GlobalSecurityData = @()
							if ($node.IE.Enabled -eq '1')
							{
								$GlobalSecurityTable += ,('SPAN-L', 'Global Security Options', '', '')
								if ($node.IE.EnhPassSecurity -eq 0)
								{
									$GlobalSecurityTable += ,('Enhanced Password Security', 'False', '', '')
								}
								else
								{
									$GlobalSecurityTable += ,('Enhanced Password Security', 'True', '', '')
									$GlobalSecurityTable += ,('Minimum Password Length', (Test-ForNull -LookupTable $null -value $node.IE.MinPassLen), '', '')
									$GlobalSecurityTable += ,('Minimum Upper Case Characters', (Test-ForNull -LookupTable $null -value $node.IE.MinPassUpper), '', '')
									$GlobalSecurityTable += ,('Minimum Lower Case Characters', (Test-ForNull -LookupTable $null -value $node.IE.MinPassLower), '', '')
									$GlobalSecurityTable += ,('Minimum Digit Characters', (Test-ForNull -LookupTable $null -value $node.IE.MinPassDigits), '', '')
									$GlobalSecurityTable += ,('Minimum Special Characters', (Test-ForNull -LookupTable $null -value $node.IE.MinPassSpecial), '', '')
									$GlobalSecurityTable += ,('Minimum Delta Previous Password', (Test-ForNull -LookupTable $null -value $node.IE.MinPassDelta), '', '')
									$GlobalSecurityTable += ,('Maximum Consecutive Characters', (Test-ForNull -LookupTable $ConsecCharsLookup -value $node.IE.MaxPassConsec), '', '')
								}
								if ($node.IE.SetPasswordLife -eq 1)
								{
									$GlobalSecurityTable += ,('Set Password Lifetime', 'True', '', '')
									$GlobalSecurityTable += ,('Maximum Password Lifetime', ($node.IE.MaxPassLife+ " days [30..365]"), '', '')
								}
								else
								{
									$GlobalSecurityTable += ,('Set Password Lifetime', 'False', '', '')
								}
								if ($node.IE.NumFailedLoginsForLockout -eq 0)
								{
									$GlobalSecurityTable += ,('Number of Failed Logins to Lockout', 'No Lockout', '', '')
								}
								else
								{
									$GlobalSecurityTable += ,('Number of Failed Logins to Lockout', $node.IE.NumFailedLoginsForLockout, '', '')
									$GlobalSecurityTable += ,('Lockout Duration', $node.IE.LockoutDuration, '', '')
								}
								if (!$SWeLite)
								{
									$GlobalSecurityTable += ,('Password Recovery', (Test-ForNull -LookupTable $TrueFalseLookup -value $node.IE.PasswordRecovery), '', '')
								}
								$GlobalSecurityTable += ,('Password Display', (Test-ForNull -LookupTable $TrueFalseLookup -value $node.IE.PasswordDisplay), '', '')
								$GlobalSecurityTable += ,('Explicit Acknowledgement of Pre-Login Info', (Test-ForNull -LookupTable $TrueFalseLookup -value $node.IE.ExpAckPreLoginInfo), '', '')
								#$GlobalSecurityData += ,('Global Security Options', '', '', $GlobalSecurityTable) - this line has moved below due to the special handling $SecurityData requires
							}
						}
					}
					#---------- TCA Configuration ------------------------
					'TCA'
					{
						$TCAObjects = $node.GetElementsByTagName('ID')
						$TCAObjectData = @()
						$TCAStatisticsCollection = @()
						$TCAStatisticsColumnTitles = @('Monitored Statistic', 'Raise Threshold', 'Clear Threshold', 'Admin State')
						if ($TCAObjects.Count -ne 0)
						{
							ForEach ($TCAObject in $TCAObjects)
							{
								if ($TCAObject.IE.classname -eq $null) { continue } # Empty / deleted entry
								if ($TCAObject.IE.classname -eq 'TCA_CONFIG')
								{
									#My 1k now displays '12': if (($TCAObject.IE.Statistic -eq '12') -and ($platform -ne 'SBC 2000')) { continue } # '12' is always in the config file but only valid in a 2k.
									if (($TCAObject.IE.Statistic -eq 1) -and ($SweLite)) { continue } # '1' is always in the config file but not valid in the SWe Lite
									if (($TCAObject.IE.Statistic -eq 4) -and ($SweLite)) { continue } # '4' is always in the config file but not valid in the SWe Lite
									$TCAObjectData = @($TCAMonitoredStatisticLookup.Get_Item($TCAObject.IE.Statistic), ($TCAObject.IE.RaiseThreshold + $TCAMonitoredStatisticValueLookup.Get_Item($TCAObject.IE.Statistic)), ($TCAObject.IE.ClearThreshold + $TCAMonitoredStatisticValueLookup.Get_Item($TCAObject.IE.Statistic)), $EnabledLookup.Get_Item($TCAObject.IE.Enabled))
									$TCAStatisticsCollection += , $TCAObjectData
								}
							}
							# Moved to the end to be correctly sequenced:
							#$SNMPData += , ('TCA Configuration', '', $TCAStatisticsColumnTitles, $TCAStatisticsCollection)
						}
					}
					
					#---------- Notification Service ------------------------
					'NotificationService'
					{
						$NotificationGroups = $node.GetElementsByTagName('Token')
						ForEach ($NotificationGroup in $NotificationGroups)
						{
							# ---- Notification Manager Config ----------
							if ($NotificationGroup.name -eq 'NotificationManager')
							{
								$NotificationData = @()
								$NotificationProfiles = $NotificationGroup.GetElementsByTagName('ID')
								if ($NotificationProfiles.Count -ne 0)
								{
									ForEach ($NotificationProfile in $NotificationProfiles)
									{
										if ($NotificationProfile.IE.classname -eq $null) { continue } # Empty / deleted entry
										if ($NotificationProfile.IE.classname -eq 'NOTIFICATION_MANAGER_IE')
										{
											$NotificationTable = @()
											$NotificationTable += ,('SPAN-L', 'Notification Manager Table', '', '')
											$NotificationDescription = (Fix-NullDescription -TableDescription $NotificationProfile.IE.Description -TableValue $NotificationProfile.value -TablePrefix 'Event #')
											$NotificationTable += ,('Description', $NotificationDescription, '' , '')
											$NotificationTable += ,('Admin State', (Test-ForNull -LookupTable $EnabledLookup -value $NotificationProfile.IE.customAdminState), '' , '')
											$NotificationTable += ,('Service Provider', (Test-ForNull -LookupTable $NotificationProviderLookup -value $NotificationProfile.IE.ServiceProvider), '' , '')
											$NotificationTable += ,('Caller ID', $NotificationProfile.IE.ProjectID, '' , '')
											$NotificationTable += ,('Client ID', $NotificationProfile.IE.ClientID, '' , '')
											$NotificationTable += ,('Secret', '****', '' , '')
											$NotificationTable += ,('Monitoring Interval', ($NotificationProfile.IE.MonitoringInterval + ' hours'), '' , '')
											$NotificationTable += ,('Events', (Test-ForNull -LookupTable $NotificationEventsLookup -value $NotificationProfile.IE.Events), '' , '')
											$E911Recipients = [regex]::replace($NotificationProfile.IE.E911RecipientsList, ',' , "`n") # Write each recipient on a new line 
											$NotificationTable += ,('E911 Recipients List', $E911Recipients, '' , '')
											$NotificationTable += ,('E911 Message', $NotificationProfile.IE.E911Message, '' , '')
										}
										$EmergSvcsData += ,('Notification Manager', $NotificationDescription, '', $NotificationTable)
									}
								}
							}
						}
					}
				}
			}

			$id = 'Finished reading data'

			#---------------------------------------------------------------------------------------------------------
			#Order everything in screen-layout sequence (as distinct from how it's arranged in the XML file):
			#---------------------------------------------------------------------------------------------------------

			$NodeHardwareData = @()
			if ($FXSCardCollection.count -ne 0) { $NodeHardwareData += , ('Ports', 'FXS Cards', $FXSCardColumnTitles, $FXSCardCollection) }
			if ($FXOLineCardCollection.count -ne 0) { $NodeHardwareData += , ('Ports', 'FXO Line Cards', $FXOLineCardColumnTitles, $FXOLineCardCollection) }
			if ($BRILineCardCollection.count -ne 0) { $NodeHardwareData += , ('Ports', 'BRI Line Cards', $BRILineCardColumnTitles, $BRILineCardCollection) }
			if ($LineCardCollection.count -ne 0) { $NodeHardwareData += , ('Ports', 'E1/T1', $LineCardColumnTitles, $LineCardCollection) }

			$NodeHardwareData += $NodeHardwarePortData
			$NodeHardwareData += $NodeHardwareLogicalInterfaceData
			if ($SWeLite)
			{
				#Skip these guys - not implemented.
			}
			else
			{
				if ($RegionSettingsDataCollection.count -ne 0) { $NodeHardwareData += ,('Bridge', 'Region Settings', $RegionSettingsColumnTitles, $RegionSettingsDataCollection) }
				if ($MSTPDataCollection.count -ne 0) { $NodeHardwareData += ,('Bridge', 'MSTP', $MSTPColumnTitles, $MSTPDataCollection) }
			}
			if ($VLANDataCollection.count -ne 0) { $NodeHardwareData += ,('Bridge', 'VLAN', $VLANColumnTitles, $VLANDataCollection) }
			$NodeHardwareData += $RelayData

			$SystemData = @()
			$SystemData += , ('Node-Level Settings', '', '', $SystemNetValues)
			if ($SWeLite)
			{
				#Skip QoE - not implemented.
			}
			else
			{
				$SystemData += , ('QoE', '', $QoEColumnTitles, $QoECollection)
			}
			$SystemData += , ('System Timing', '', $TimingConfigColumnTitles, $TimingConfigCollection)
			$SystemData += , ('System Companding Law', '', $CompandingLawConfigColumnTitles, $CompandingLawConfigCollection)
			if (!$SWeLite)
			{
				#The SweLite has no Port Licences table
				if ($PortLicenceCollection.count -ne 0) { $SystemData += , ('Licensing', 'Current Licenses - Port Licenses', $PortLicenceColumnTitles, $PortLicenceCollection) }
			}
			if ($FeatureLicenceCollection.count -ne 0) { $SystemData += , ('Licensing', 'Current Licenses - Feature Licenses', $FeatureLicenceColumnTitles, $FeatureLicenceCollection) }

			$IPRouteData = @()
			$IPRouteData +=  $DNSData
			# IP (v4)
				$IPRouteData += $StaticRouteData
				#$IPRouteData +=  # Routing Table
				#$IPRouteData +=  # Router Instances
				$IPRouteData +=  $AllACLData
				#$IPRouteData +=  # NAT
			# IPv6
				#$IPRouteData =  # Static Routes
				#$IPRouteData =  # Routing Table
				$IPRouteData +=  $AllACLv6Data
			$IPRouteData +=  $DHCPData
			#$IPRouteData +=  # IP Sec
			$IPRouteData += $NetworkMonitoringData
			$IPRouteData += $CACData

			$MediaData = @()
			$MediaData += $MediaSystemConfigData
			$MediaData += $MediaProfilesData
			$MediaData += $MediaFaxData
			$MediaData += $SDESProfilesData
			$MediaData += $DTLSProfilesData
			$MediaData += $MediaListData

			#'Security' needs special handling to piece the array together in the right order: it doesn't flow in the config file the same way as it is on-screen
			if ($GlobalSecurityTable.count -ne $null) { $SecurityData += , ('Users', 'Global Security Options', '', $GlobalSecurityTable) }
			if ($ADUserGroupCollection.count -ne $null) { $SecurityData += , ('Users', 'Remote Auth Permissions / AD User Group', $ADUserGroupColumnTitles, $ADUserGroupCollection) }
			$SecurityData += , ('Users', 'Remote Auth Permissions / RADIUS User Class', $RadiusClassColumnTitles, $RadiusClassCollection)
			$SecurityData +=   ($CertificateCollectionMy)
			$SecurityData +=   ($CertificateCollectionRoot)
			$SecurityData +=   ($TLSCollection)
			$SecurityData += , ('Bad Actors', '', $BadActorColumnTitles, $BadActorCollection)

			#Logging Data needs $PortMirrorData glued to the end so they show in the correct sequence:
			$LoggingData += $PortMirrorData

			if ($SNMPManagersCollection.count  -ne 0) { $SNMPData += , ('SNMP Management Addresses', '', $SNMPManagerColumnTitles, $SNMPManagersCollection) }
			if ($SNMPEventsCollection.count	-ne 0) { $SNMPData += , ('Alarms/Events Config', '', $SNMPEventsColumnTitles, $SNMPEventsCollection) }
			if ($TCAStatisticsCollection.count -ne 0) { $SNMPData += , ('TCA Configuration', '', $TCAStatisticsColumnTitles, $TCAStatisticsCollection) }

			$AllSipData = @()
			$AllSipData += $AllSipLocalRegistrarData
			$AllSipData += $AllSipLocalPassthroughData
			$AllSipData += $AllSIPProfileData
			$AllSipData += $AllSipServerData
			$AllSipData += $AllSIPTrunkGroupData
			$AllSipData += $AllSIPNATPrefixData
			$AllSipData += $AllSipRemoteAuthData
			$AllSipData += $AllSipContactRegistrantData
			$AllSipData += $AllSipMsgRuleData
			$AllSipData += $SIPCondRuleData
			$AllSipData += $NodeLevelSIPData
			$AllSipData += $AllSIPRecData

			$AllCallRouting = @()
			$AllCallRouting += $TransformationData
			$AllCallRouting += $TODData
			$AllCallRouting += $RouteData
			$AllCallRouting += $ActionConfigData

			$Sections = @()
			if ($DoAll -or $DoCalls)  { $Sections += ,('Call Routing', $AllCallRouting) }

			if ($DoAll -or $DoSig)	{ $Sections += ,('Signaling Groups', $SGData) }
			if ($DoAll -or $DoSig)	{ $Sections += ,('Linked Signaling Groups', $LinkedSGData) }
			if ($SWeLite)
			{
				if ($DoAll -or $DoIP)	 { $Sections += ,('Node Interfaces', $NodeHardwareData ) }
			}
			else
			{
				if ($DoAll -or $DoIP)	 { $Sections += ,('Networking Interfaces', $NodeHardwareData ) }
			}
			if ($DoAll -or $DoSystem)	{ $Sections += ,('Application Solution Module', $SBAData ) }
			if ($DoAll -or $DoSystem)	{ $Sections += ,('System', $SystemData) }
			if ($DoAll -or $DoSystem)	{ $Sections += ,('Auth and Directory Services', $ADData) }
			if ($DoAll -or $DoIP)		{ $Sections += ,('Protocols', $IPRouteData) }
			if ($DoAll -or $DoSIP)		{ $Sections += ,('SIP', $AllSipData) }
			if ($DoAll -or $DoSig)		{ $Sections += ,('CAS', $CASData) }
			if ($DoAll -or $DoMisc)		{ $Sections += ,('Security', $SecurityData) }
			if ($DoAll -or $DoMisc)		{ $Sections += ,('Media', $MediaData) }
			if ($DoAll -or $DoMisc)		{ $Sections += ,('Tone Tables', $ToneTableData) }
			if ($DoAll -or $DoSig)		{ $Sections += ,('Telephony Mapping Tables', $TMTData) }
			if ($DoAll -or $DoMaint)	{ $Sections += ,('SNMP/Alarms', $SNMPData) }
			if ($DoAll -or $DoMaint)	{ $Sections += ,('Logging Configuration', $LoggingData) }
			if ($DoAll -or $DoCalls)	{ $Sections += ,('Emergency Services', $EmergSvcsData) }
			#if ($DoAll -or $DoCalls)	{ $Sections += ,('Emergency Services', $NotificationData) }

			$sectionCounter = 0
			foreach ($section in $Sections)
			{
				[string]$H1title , $SectionData = $section

				$sectionCounter++
				$Progress1Complete = 10 + (($sectionCounter/$sections.Count)*80) #I'm fudging a little: we assume we're 10% by the time we get to here, and 90% complete when we're done writing all sections.
				write-progress -id 1 -Activity ("Writing to '{0}'" -f ($RawOutputFileName)) -Status $H1title -PercentComplete ($Progress1Complete)

				if ($SectionData.Count -eq 0)  { continue } #Nothing to write - move on.
				if ($null -eq $SectionData )   { continue } #Nothing to write - move on.
				# The line above was added in 2.4 to trap a problem peculiar to $PSv2.
				# Changed in v5.0.1 to reverse the test: apparently ($sectionData -eq $null) returns true if ANY element in the array is null - not if the whole thing is.

				#Write the section Heading
				$selection.Style= $wdStyleHeading1 # 'Heading 1'
				$selection.TypeText($H1title)
				$selection.ParagraphFormat.KeepWithNext = -1
				$selection.TypeParagraph()
				write-verbose -message ('Writing: {0}' -f ($H1title))

				$LastH2 = 'Greiginsydney.com' # ;-)
				foreach ($SectionEntry in $SectionData)
				{
					[string]$subHeadText, [string]$tableHeading, [array]$columnHeaderRow, [array]$tableContent = $SectionEntry
					#if ($tableContent.Count -eq 0) { continue } #Nothing to write - move on. #Removed in 7.0.0B to now show empty tables/objects in the document
					if ($null -eq $tableContent  ) { continue } #Nothing to write - move on.
					# The line above was added in 2.6 to trap a problem peculiar to $PSv2.
					# Changed in v5.0.1 to reverse the test: apparently ($sectionData -eq $null) returns true if ANY element in the array is null - not if the whole thing is.
					if ($subHeadText -eq '')
					{
						#Then there's no sub-heading. This is a stand-alone table.
						$ProgressBar2Text = $H1title
					}
					else
					{
						#Skip sub-head if it's the same as last time
						if ($LastH2 -ne $subHeadText)
						{
							write-verbose -message ('Writing: - {0}' -f ($subHeadText))
							$selection.Style= $wdStyleHeading2 # 'Heading 2'
							$selection.TypeText($subHeadText)
							$selection.ParagraphFormat.KeepWithNext = -1
							$selection.TypeParagraph()
							$LastH2 = $subHeadText
							$ProgressBar2Text = $subHeadText
						}
					}
					if ($tableHeading -ne '')
					{
						write-verbose -message ('Writing:  - {0}' -f ($tableHeading))
						$selection.Style= $wdStyleHeading3 #'Heading 3'
						$selection.TypeText($tableHeading)
						$selection.ParagraphFormat.KeepWithNext = -1
						$selection.TypeParagraph()
						$ProgressBar2Text = $tableHeading
					}
					if ($columnHeaderRow.Length -le 1)
					{
						#ColumnHeaderRow isn't an array, so it's a vertical table.
						#All we're doing here is changing the text we pass, which becomes the title for the progress bar.
						Write-SectionVertically -progressBarTitle $ProgressBar2Text -values $tableContent
					}
					else
					{
						Write-Section -progressBarTitle $ProgressBar2Text -headers $columnHeaderRow -values $tableContent
					}
				}
				write-verbose -message ('Writing: {0} - complete' -f ($H1title))
			}

			#Add closing text:
			$selection.TypeParagraph()
			$selection.Paragraphs.Alignment = $wdAlignParagraphCenter
			$selection.TypeText('---------- End of Document ----------')


			if ($RedactIP)
			{
				$id = 'De-identifying: removing IPs'
				#--------------------------------
				# Find/Replace to remove IP addresses
				#--------------------------------
				$IPAsRegex = '(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)'
				$ReplaceIP = 1	# This is the dummy suffix we add to each unique IP address in the file
				$range = $doc.content
				while (1)
				{
					$null = $range.movestart()
					If ($range.Text -match $IPAsRegex)
					{
						$found = $Matches[0]
						$find = $selection.Find.Execute($found, $false, $false, $false, $false, $false, $false, $false, $false, ('<IPAddress-' + $ReplaceIP.ToString() + '>'), $wdReplaceAll)
						$ReplaceIP ++
					}
					else
					{
						break
					}
				}
			}

			$id = 'Global paragraph formatting'
			#--------------------------------
			# Apply global paragraph formatting - this tightens up the tables, removing leading or trailing paragraph spacing. Apply some if you want to space them out a bit more.
			#--------------------------------
			$paragraphs = $doc.paragraphs
			$paragraphs.SpaceBeforeAuto = $False
			$paragraphs.SpaceBefore = '0'
			$paragraphs.SpaceAfterAuto = $False
			$paragraphs.SpaceAfter = '0'
			#--------------------------------

			$id = 'Global table formatting'
			#--------------------------------
			# Apply global formatting to all the tables
			#--------------------------------
			write-progress -id 1 -Activity 'Finishing Up' -Status 'Applying global table formatting' -PercentComplete (92) # (Near enough...)
			write-verbose -message 'Applying global table formatting'
			$TablesCount = ($doc.Tables).Count
			$ThisTable = 0
			foreach ($table in $doc.Tables)
			{
				$ThisTable++
				$progressPercent2 = ($ThisTable/$TablesCount*100)
				If ($progressPercent2 -gt 100) {$progressPercent2 = 100 } #Trap any accidental percentage overrun
				write-progress -id 2 -parentid 1 -Activity 'Table' -Status 'Formatting' -PercentComplete $progressPercent2

				$selection = $table.Rows
				$selection.AllowBreakAcrossPages = $False

				if ($WdTableStyle)
				{
					$table.Style = $WdTableStyle
				}
				$table.AutoFitBehavior(2) # http://msdn.microsoft.com/en-us/library/office/bb213671(v=office.12).aspx
			}
			write-progress -id 2 -parentid 1 -Activity 'Table' -Status 'Formatting' -PercentComplete (100) -Completed

			$id = 'Setting the footer'
			#--------------------------------
			# add the footer ----------------
			#--------------------------------
			#set the footer
			write-progress -id 1 -Activity 'Finishing Up' -Status 'Setting the footer' -PercentComplete (94) # (Near enough...)
			write-verbose -message 'Set the footer'
			[string]$footertext=('{0}. Created by {1}' -f ($RawOutputFileName), $env:username)

			#get the footer
			write-verbose -message 'Get the footer and format font'
			$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekPrimaryFooter
			#get the footer and format font
			$footers=$doc.Sections.Last.Footers
			foreach ($footer in $footers)
			{
				if ($footer.exists)
				{
					if (!$WdTemplate)
					{
						$footer.range.Font.name='Calibri'
						$footer.range.Font.size=8
						$footer.range.Font.Italic=$True
						$footer.range.Font.Bold=$True
					}
				}
			}
			write-verbose -message 'Footer text'
			$selection=$word.Selection
			$selection.HeaderFooter.Range.Text=$footerText

			#add page numbering
			write-verbose -Message 'Add page numbering'
			$null = $selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight)

			#return focus to main document
			write-verbose -message 'Return focus to main document'
			$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument


			#--------------------------------
			# update the TOC ----------------
			#--------------------------------
			write-progress -id 1 -Activity 'Finishing Up' -Status 'Updating the Table of Contents' -PercentComplete (96) # (Near enough...)
			foreach ($TOC in $doc.TablesOfContents)
			{
				#$TOC.UseHyperlinks = -1
				$TOC.Update()
				$TOC.UpdatePageNumbers()
			}
			write-progress -id 1 -Activity 'Finishing Up' -Status 'Updated the Table of Contents' -PercentComplete (96) # (Near enough...)
		}
		catch
		{
			$ErrLine = $_.ScriptStackTrace | select-string -pattern '.ps1: line (\d*)' | ForEach-Object{($_.Matches)} | ForEach-Object {$_.Groups[1].Value}
			if ($_.Exception -match 'Key cannot be null')
			{
				write-host ('Aborted with unhandled value error at line {0}' -f ($Errline)) -ForegroundColor Red -BackgroundColor Black
				if ($id -ne '') { write-host ('$id = {0}' -f ($id)) -ForegroundColor Red -BackgroundColor Black }
			}
			else
			{
				write-host ("Aborted '{0}' with fatal error @ line {1}:" -f ($InputFile), ($Errline)) -ForegroundColor Red -BackgroundColor Black
				Dump-ErrToScreen
			}
		}
		finally
		{
			#This whole section exists to both ensure we close/exit Word if there's an error or the user aborts (^c),
			# or if you try to re-run the script while the output file is already open/in use.
			try
			{
				#--------------------------------
				# save ... ----------------------
				#--------------------------------
				if ($doc)
				{
					#(It might not be open if we've dropped through to here after a failure up the top ('cause you left the document open in Word!)
					if ($Error[0] -like '*is already open elsewhere*')
					{
						#We're here through error - this is not a normal exit (or a user-abort).
						#Don't try to Save the file - that's why we're here in the first place!
					}
					else
					{
						write-verbose -message 'Save document'
						$doc.Save()
						if ($MakePDF -eq $True)
						{
							write-progress -id 1 -Activity 'Finishing Up' -Status 'Saving as PDF' -PercentComplete (98) # (Near enough...)
							$PdfOutputFile = ([IO.Path]::ChangeExtension($OutputFile, 'pdf'))
							write-verbose -message ('Saving as PDF - "{0}"' -f ($PdfOutputFile))
							$wdCreateBookmarks = [Microsoft.Office.Interop.Word.WdExportCreateBookmarks]::wdExportCreateHeadingBookmarks # 1
							$WdFixedFormatExtClassPtr = [Type]::Missing
							# OK, this is a repeat of the same ugly kludge from earlier, but if it doesn't want to Save using '[ref]', let's try again without it:
							try
							{
								$doc.ExportAsFixedFormat([ref]$PdfOutputFile, [ref]17, $false, 1, 0, 0, 0, 0, $true, $true, 1)#, $true, $true, $false, [ref]$WdFixedFormatExtClassPtr)
							}
							catch
							{
								$doc.ExportAsFixedFormat($PdfOutputFile, 17, $false, 1, 0, 0, 0, 0, $true, $true, 1)#, $true, $true, $false, [ref]$WdFixedFormatExtClassPtr)
							}
						}
					}
				}
			}
			catch
			{
				write-warning -message 'Failed to save DOC and/or create PDF:' # If we can't save, too bad.
				if ($_.Exception -match 'This file is in use')
				{
					write-warning -message 'The file is in use by another application or user'
				}
				else
				{
					Dump-ErrToScreen
				}
			}
			try
			{
				#--------------------------------
				# ... and exit ------------------
				#--------------------------------
				try
				{
					if ($doc) { $doc.Close($wdDoNotSaveChanges) }
				}
				catch [Management.Automation.MethodException] # Oh Windows 7 / P$v2, how you taunt me!
				{
					if ($doc) { $doc.Close([ref]$wdDoNotSaveChanges) }
				}
				if ($word)
				{
					write-verbose -message 'Exit word'
					$word.quit()
					# Stop Winword Process
					$rc = [Runtime.Interopservices.Marshal]::ReleaseComObject($word)
				}
			}
			catch
			{
				write-warning -message 'Error closing Word'
				Dump-ErrToScreen
			}
			write-verbose -message 'Done!'
		}
	}

	#--------------------------------
	# END FUNCTIONS -----------------
	#--------------------------------



	if ($skipupdatecheck)
	{
		write-verbose -message 'Skipping update check'
	}
	else
	{
		write-progress -id 1 -Activity 'Initialising' -Status 'Performing update check' -PercentComplete (2)
		Get-UpdateInfo -title 'Convert-RibbonSbcConfigToWord.ps1'
		write-progress -id 1 -Activity 'Initialising' -Status 'Back from performing update check' -PercentComplete (2)
	}

	write-progress -id 1 -Activity 'Initialising' -Status 'Initialising' -PercentComplete (1)
	#If the user only provided a filename, add the script's path for an absolute reference:
	$scriptpath = $MyInvocation.MyCommand.Path
	$dir = Split-Path -Path $scriptpath

	$PSVersion = $PSVersionTable.PsVersion.Major 	#As of v5.0.1 I'm going to simply bypass code that I can't readily port backwards to v2.



} # END "Begin"


process
{

	#(Re-)initialise all the lookup tables for each loop:
	$PortToIPAddressLookup = @{'' = 'Auto'}			#The names of the NIC ports & their IP addresses	Referenced by: SIP Signaling Gps, DHCP Pools & ACLs
	$ACLTableLookup	= @{'0' = 'None'}				#The names of the ACL tables 			Referenced by: Ports / Logical Interfaces
	$ACLv6TableLookup	= @{'0' = 'None'}			#The names of the IPv6 ACL tables		Referenced by: Ports / Logical Interfaces
	#$TranslationTableLookup = @{}					#The names of the Translation Tables	Referenced by: Call Routes					- moved in v6.0 with 'FixNullAndDuplicateDescriptions'
	#$TransformationTableLookup = @{'0' = 'None'}	#The names of the Transformation tables	Referenced by: Call Routes & Action Sets		- moved in v6.0 with 'FixNullAndDuplicateDescriptions'
	$SgTableLookup = @{}							#The names of the Signaling Groups		Referenced by: Call Routes, & Signaling Groups (from v9.0.0+)
	$CallRoutingTableLookup = @{}					#The names of the call routing tables	Referenced by: Action Config & Signaling Groups
	$ActionSetLookup = @{'0' = 'None'} 				#The names of the Action Sets			Referenced by: Signaling Gps
	$ActionSetConfigLookup = @{'0'='Continue'}		#The names of the Action Configs		Referenced by: Action Set Tables
	$ToneTableLookup = @{}							#The names of the Tone Tables			Referenced by: Signaling Gps
	$SipProfileIdLookup = @{}						#The names of the SIP Profiles			Referenced by: SIP Signaling Gps
	$SipCredentialsTableLookup = @{'0' = 'None';}	#The names of the SIP Credential Tables	Referenced by: SIP Servers
	$SipRegistrationTableLookup = @{'0' = 'None';}	#The names of the SIP Registration Tables		Referenced by: SIP Servers
	$SIPAuthorisationTableLookup = @{'0' = 'None';}	#The names of the SIP Authorisation Tables		Referenced by: SIP Sig Gps
	$SIPMessageRuleLookup = @{'0' = 'None';} 		#The names of the SIP Message Manipulation Rule Tables	Referenced by: SIP Servers
	$SIPMessageRuleElementLookup = @{'0' = 'None';}	#The names of the SIP Msg Rule Element Descriptors	Referenced by: SIP Message Rule Tables
	$SIPCondRuleLookup = @{}						#The names of the SIP Condition Rule Tables		Referenced SIP Message Rule Tables
	$TlsProfileIdLookup = @{'0' = 'N/A';}			#The names of the TLS Profiles			Referenced by: SIP Servers
	$CertificateLookup = @{'-1000' = '--- No Tables ---';}		#The names of the Certificates	Referenced by: TLS Profiles
	$SDESMediaCryptoProfileLookup = @{'0' = 'None';}	#The names of the SDES Crypto Profiles		 Referenced by: Media List Profiles
	$DTLSMediaCryptoProfileLookup = @{'0' = 'None';}	#The names of the DTLS Crypto Profiles		 Referenced by: Media List Profiles
	$SIPServerTablesLookup = @{'0' = 'None';}		#The names of the SIP server tables		Referenced by: SIP Signaling Gps, Node-Level SIP Settings & SIP Recording
	$SIPRegistrarsLookup = @{}						#The names of the SIP registrars		Referenced by: SIP Signaling Gps
	$VoiceFaxProfilesLookup = @{}					#The names of the Media Types			Referenced by: Media List Profiles
	$MediaListProfileLookup = @{'0' = 'None';}		#The names of the Media List Profiles	Referenced by: SIP Signaling Gps & Call Routes
	$CASSignalingProfileLookup = @{}				#The names of the CAS Signaling Profiles	Referenced by: CAS Signaling Gps
	$CASSupplementaryProfileLookup = @{'0' = 'None';}	#The names of the CAS Signaling Profiles	Referenced by: CAS Signaling Gps
	$MsgTranslationTablesLookup = @{'0' = 'None';}	#The names of the Msg Translation Tables	Referenced by: Call Routes
	$RerouteTablesLookup = @{'0' = 'None';}			#The names of the Reroute Tables			Referenced by: Call Routes
	$SIPtoQ850TableLookup  = @{'0' = 'Default (RFC4497)';}	#The names of the SIPtoQ850 Tables		Referenced by: SIP Sig Gps
	$Q850toSIPTableLookup  = @{'0' = 'Default (RFC4497)';}	#The names of the Q850toSIP Tables		Referenced by: SIP Sig Gps
	$RadiusServerLookup = @{'0' = 'None';}			#The names of the RADIUS Servers		Referenced by: Auth & Dir Services / RADIUS / Configuration
	$NetworkVLANLookup = @{}						#The names of the VLANs					Referenced by: Ports / Ethernet interfaces
	$MSTPInstanceNameLookup = @{}					#The names of the STP instances			Referenced by: Ports / Bridge / VLAN
	$MSTPInstanceIdLookup = @{}						#The IDs of the STP instances			Referenced by: Ports / Bridge / VLAN
	$SIPNATPrefixesLookup = @{'0' = 'None';}		#The names of the SIP NAT Prefix Tables	Referenced by: SIP Signaling Gps
	$LogServerLookup = @{'0' = 'Local Logs';}		#The names of the Log Servers			Referenced by: Logging Configuration / Subsystems
	$DHCPPoolLookup = @{}							#The names of the DHCP Pools (scopes)	Referenced by: DHCP Static Clients
	$TODTablesLookup = @{'0' = 'None';}				#The names of the Time of Day Tables	Referenced by: Call Routes
	$CACProfileLookup = @{}							#The names of the Skype/Lync CAC profiles	Referenced by: Network Monitoring / Link Monitors
	$IPSecTunnelLookup = @{}						#The names of the IP Sec tunnel tables	Referenced by: Network Monitoring / Link Monitors


	#write-warning "Inputfile = $($inputfile)"	- I'll keep these here until begin/process/end is bedded down some more
	#write-warning "FullName  = $($FullName)"

	$Params = @{
		Do = $Do
		OutputFile = $OutputFile
		NodeID = $NodeID
		HardwareID = $HardwareID
		TitleColour = $TitleColour
		LabelColour = $LabelColour
		WdTemplate = $WdTemplate
		WdTableStyle = $WdTableStyle
	}

	#Add the switches if they were provided:
	if ($IncludeNodeInfo.IsPresent) { $Params.Add("IncludeNodeInfo", $true) }
	if ($MakePDF.IsPresent) { $Params.Add("MakePDF", $true) }
	if ($Landscape.IsPresent) { $Params.Add("Landscape", $true) }
	if ($RedactIP.IsPresent) { $Params.Add("RedactIP", $true) }
	if ($SonusGreen.IsPresent) { $Params.Add("SonusGreen", $true) }
	if ($SkipWrite.IsPresent) { $Params.Add("SkipWrite", $true) }

	If ([string]::IsNullOrEmpty($FullName))
	{
		$Params.Add("InputFile", $InputFile)
		#write-host "Passing Inputfile = $($inputfile)"
	}
	else
	{
		$Params.Add("InputFile", $FullName)
		#write-host "Passing FullName = $($FullName)"
	}
	Do-Convert @Params
}

end
{
	#write-warning "Sayonara"
}



# With thanks to:
# http://jamesmccaffrey.wordpress.com/2007/12/02/parsing-xml-files-with-powershell/
# TOC: http://www.edrawsoft.com/automating-word-vb.php
# http://carlwebster.com/documenting-a-citrix-xenapp-6-5-farm-with-microsoft-powershell-and-word-version-3/
# http://stackoverflow.com/questions/14507815/powershell-word-2007-2010-2013-add-a-caption-to-table-and-move-table-3-tabs-to-r
# http://mypowershell.webnode.sk/news/create-word-document-with-multiple-tables-from-powershell/
# http://dilutedthoughts.com/dilutedthoughts/2012/12/19/using-powershell-to-create-a-custom-word-document
# Save as PDF: http://blog.coolorange.com/2012/04/20/export-word-to-pdf-using-powershell/
# Write-Progress: http://www.hanselman.com/blog/ProgressBarsInPowerShell.aspx
# Portrait vs Landscape: http://blogs.technet.com/b/heyscriptingguy/archive/2006/08/31/how-can-i-set-the-document-orientation-in-microsoft-word-to-landscape.aspx
# Extract from zip: http://www.howtogeek.com/tips/how-to-extract-zip-files-using-powershell/
# Search/Replace (for RedactIP) https://www.safaribooksonline.com/library/view/regular-expressions-cookbook/9780596802837/ch07s16.html
# Search/Replace (for RedactIP) http://stackoverflow.com/questions/27169043/powershell-search-matching-string-in-word-document
# Search/Replace (for RedactIP) https://gallery.technet.microsoft.com/office/7c463ad7-0eed-4792-8236-38434f891e0e
# Testing for null in PowerShell: http://stackoverflow.com/questions/29614171/testing-for-null-in-powershell-why-does-testing-an-empty-array-behave-different
# Custom sort: (Same answer - who came first?) https://powershellone.wordpress.com/2015/07/30/sort-data-using-a-custom-list-in-powershell/
# Custom sort: (Same answer - who came first?) https://gallery.technet.microsoft.com/scriptcenter/Sort-With-Custom-List-07b1d93a
# Sorting multidimensional arrays: https://www.sherweb.com/blog/fun-with-powershell-sorting-multidimensional-arrays-and-arraylists/
# Title case: https://techibee.com/powershell/powershell-how-to-change-text-to-title-casefirst-letter-to-upper-case-and-remaining-lower-case/1235

# Interesting comment here: http://msdn.microsoft.com/en-us/library/bb221597.aspx
# "Another thing to be aware of: Use backslashes in the file path only, not forward slashes."


#Code signing certificate kindly provided by Digicert:
