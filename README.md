# Convert-RibbonSbcConfigToWord.ps1
The name’s a bit of a mouthful, but “Convert-RibbonSbcConfigToWord.ps1” takes the backup file from your Sonus/Ribbon SBC 1000/2000/SWe Lite gateway and creates a new Word document, with all of the important(?) configuration information captured in tables.

<p>&nbsp;</p>
<p><span style="font-size: small;">There's more information on the blog:</span></p>
<p><span style="font-size: medium;">https://greiginsydney.com/uxbuilder</span></p>
<hr />
<p>&nbsp;</p>
<p>The name&rsquo;s a bit of a mouthful, but &ldquo;Convert-RibbonSbcConfigToWord.ps1&rdquo; takes the backup file from your Sonus/Ribbon SBC 1000/2000/SWeLite gateway and creates a new Word document, with all of the important(?) configuration information captured  in tables.&nbsp;</p>
<p>It started life as a way to save the tedium of screen-scraping lots of fixed frames for my as-built documents, but it quickly became apparent that it would also make a useful tool for the offline review of a gateway&rsquo;s config (although it ain&rsquo;t  no &ldquo;UxBuilder&rdquo;).</p>
<p>&nbsp;</p>
<h2><span style="color: #ff0000;">Convert your backups from this:</span></h2>

<img src="https://user-images.githubusercontent.com/11004787/81054574-64b12f80-8f0a-11ea-8a57-60fef112945a.png" alt="" width="600" />

<h2><span style="color: #ff0000;">&hellip; to this:</span></h2>

<img src="https://user-images.githubusercontent.com/11004787/81054622-7abef000-8f0a-11ea-9148-28a696bd805d.png" alt="" width="600" />

<h3>Features</h3>
<ul>
<li>Decodes gateway versions up to 8.1.0 &amp; SWe Lite. </li>
<li>Converts and saves the config to a new Word document, and optionally makes a PDF of it too. </li>
<li>Doesn&rsquo;t need any Lync PowerShell modules &ndash; it&rsquo;s fine on your Win7-Win10 machine (which needs to have Word installed). </li>
<li>Tested on Word 2010, 2013 &amp; 2016. </li>
<li>Reasonably well debugged throughout (although time will tell of course!). </li>
<li>It&rsquo;s easy enough to turn off an unwanted section if you&rsquo;re only interested in a specific subset of the functionality, or are adapting to suit your environment. </li>
<li>Use the "RedactIP" switch to strip all IP addresses from the resulting file. </li>
<li>A &ldquo;-SkipWrite&rdquo; switch skips the relatively slow process of writing to Word. I added this for debugging and have left it in the released version to help anyone refining or enhancing the script. </li>
<li>From v2 the script has been signed, so no longer requires you to drop your ExecutionPolicy to &ldquo;Unrestricted&rdquo;. This also means you can be confident the code&rsquo;s not changed since it left here. </li>
<li>Uses your existing Word document template for the formatting of the table headings and table of contents. </li>
</ul>
<h3>Limitations</h3>
<ul>
<li>It&rsquo;s not the fastest to run! Stuffing values into Word is a relatively slow process. With all options enabled allow up to 10 minutes for your average gateway. I&rsquo;ve stolen Scott Hanselman&rsquo;s "<a href="http://www.hanselman.com/blog/ProgressBarsInPowerShell.aspx" target="_blank">Progress  Bars in PowerShell</a>&rdquo; to help the watched pot boil. </li>
<li>v8.0 captures even more of the gateway&rsquo;s config than before, however it still doesn&rsquo;t capture EVERYTHING. There&rsquo;s an enormous amount in there, and it&rsquo;s growing with each release! If you find I&rsquo;ve not captured a particularly  valuable element of the config, feel free to revise it to suit your needs &ndash; and send me the new bits to incorporate, please. If you&rsquo;re not feeling up to it, send me a sample config file and some screen-grabs of the bit you&rsquo;d like to see added. </li>
<li>There are MANY hash tables throughout. (This is where an integer is stored in the config file, but it&rsquo;s decoded to text on-screen). I&rsquo;ve put a lot of effort into capturing these, but there are going to be examples where I&rsquo;ve missed one.  Please send me screen-grabs and sample config files and I&rsquo;ll update the script. </li>
<li>I&rsquo;m largely letting Word control the width of the table columns, and you&rsquo;ll find many instances where it&rsquo;s done a poor job. You may need to edit every document the script produces to resolve this. </li>
</ul>
<h3>Instructions</h3>
<ol>
<li>Backup your gateway, saving the config file somewhere accessible. Older gateway versions (pre 2.2.1) will save to an archive called &ldquo;backup.tar.gz&rdquo;, while newer versions save as &ldquo;SBC_Config_&lt;hostname&gt;_&lt;version&gt;_&lt;date&gt;.tar&rdquo; </li>
<li>Older backups you&rsquo;ll need to unzip <em>twice</em> to recover the &ldquo;symphonyconfig.xml&rdquo; document we need, whilst newer ones only need one unzip to reveal their goodness <br /> &gt;&gt; If you have <a href="http://www.7-zip.org/" target="_blank">7-zip</a> installed you no longer need to unzip &ldquo;SBC_Config_&lt;hostname&gt;_&lt;version&gt;_&lt;date&gt;.tar&rdquo; - just feed the .tar file as the "-inputfile" </li>
<li>Open a PowerShell window </li>
<li>Now feed the backup into the script! Either place the config file in the same path as the script, give the full path in the command line, or a relative path:
<pre>PS H:\&gt; Convert-RibbonSbcConfigToWord.ps1 -InputFile "symphonyconfig.xml" -OutputFile "CustomerSBC-config.docx"</pre>
If you don&rsquo;t specify an InputFile, the script goes looking for "symphonyconfig.xml" in its directory. If you don&rsquo;t specify an OutputFile, it will re-use the name of the InputFile and save as &ldquo;.docx&rdquo;, overwriting any existing file of  the same name in that directory. If you add the &ldquo;-MakePdf&rdquo; switch, it will save the same output filename as a .pdf as well </li>
</ol>
<p>You can easily batch it:</p>
<pre>PS H:\&gt; Get-ChildItem "d:\path\*.xml" -recurse | .\Convert-RibbonSbcConfigToWord.ps1 -IncludeNodeInfo -MakePdf -SkipUpdateCheck</pre>
<pre>PS H:\&gt; Get-ChildItem "d:\path\*.tar" -recurse | .\Convert-RibbonSbcConfigToWord.ps1 -MakePdf -SkipUpdateCheck</pre>
&nbsp;
<h3>Revision History</h3>
<h4>(Read about older versions <a href="https://greiginsydney.com/uxbuilder" target="_blank"> on my blog</a>)</h4>
<h4>v8.1.0B 18th March 2020</h4>
<ul>
<li>Fixed bug where TT's were no longer arranged alphabetically - broken with the new layout introduced in v8.0.0. (Tks Mike.) </li>
</ul>
<h4>v8.1.0 14th March 2020</h4>
<ul>
<li>Restructured to support batch processing direct from the pipeline. Added new batch example:<br /> gci *.tar | .\Convert-RibbonSbcConfigToWord.ps1 -SkipUpdateCheck -MakePDF </li>
<li>Added new bits in 8.1.0: 
<ul>
<li>&ldquo;Country Code&rdquo; in System / Node Level Settings </li>
<li>&ldquo;Bad Actors&rdquo; table under Settings / Security </li>
<li>&ldquo;Explicit Acknowledgement of Pre-Login Info&rdquo; in Security / Users / Global Security Options </li>
<li>&ldquo;Certificate&rdquo; in Security / TLS Profiles </li>
<li>&ldquo;SBC Supplementary Certificates&rdquo; in Security / SBC Certificates </li>
<li>&ldquo;SILK&rdquo; in System / Licensing / Current Licenses </li>
<li>&ldquo;DC Enabled&rdquo; in Auth and Directory Services / Acive Directory / Domain Controllers </li>
</ul>
</li>
<li>Fixed bugs: 
<ul>
<li>&ldquo;Lockout duration&rdquo; was always showing under Security / Users / Global Security Options, even if Number of Failed Logins was set to &ldquo;No lockout&rdquo; </li>
<li>Was not correctly handling &ldquo;SupportsShouldProcess&rdquo;. As resolving this will add a confirm prompt for each iteration (potentially breaking existing workflow), I&rsquo;ve opted to disable it </li>
<li>Under System / Node-Level Settings, &lsquo;Power LED&rsquo; was misspelt </li>
</ul>
</li>
<li>Updated label on SIP Sig Sp &lsquo;Outbound Proxy&rsquo; to now show as &lsquo;Outbound Proxy IP/FQDN&rsquo; </li>
</ul>
<h4>v8.0.3 * Not released.</h4>
<ul>
<li>Added display of settings when Security / Users / Global Security Options / Enhanced Password Security = Yes </li>
</ul>
<h4>v8.0.2 * Not released.</h4>
<ul>
<li>No new functionality, no bugs unearthed. </li>
</ul>
<h4>v8.0.1 - 30th January 2019</h4>
<ul>
<li>Added new bits in 8.0.1: 
<ul>
<li>Changed ISDN Sig Gp for new 'Until Timer Expiry' option in 'Play Inband Message Post-Disconnect' </li>
<li>Added new 'Convert DISC PI=1,8 to Progress' to ISDN Sig Gp </li>
</ul>
</li>
<li>Fixed bugs: 
<ul>
<li>If the SBC has only 1 Transformation Table it would not display. (Bug in the alpha sort introduced in v7.0.0B) </li>
<li>Corrected several bugs in display of 'DTLS-SRTP Profiles' </li>
<li>Corrected bug in Media Lists where the 'DTLS-SRTP Profile' value wasn't correctly reported </li>
<li>Certs was truncating parameters if one happened to be a URL (as I was splitting on '/') </li>
<li>Certs wasn't populating the issuer's details (the RHS of the table) </li>
</ul>
</li>
<li>Created new $SipProfileCntctHdrLookup &amp; edited values in $SipProfileFrmHdrLookup for new on-screen text </li>
<li>Added "[No CN]" text if a cert (so far GoDaddy roots) has no CN </li>
</ul>
<h4>v8.0.0B - 16th November 2018</h4>
<ul>
<li>Updated label and display of SIP Sig Gp / Media Info / Supported Audio(/Fax) Modes to match current display syntax: Lite and 1k2k </li>
<li>Fixed bug: 
<ul>
<li>The logic for displaying SIP Sig Gp ICE Mode was faulty. (Mitsu-San yet again on the ball!) </li>
</ul>
</li>
</ul>
<h4>v8.0.0 - 26th September 2018</h4>
<ul>
<li>Updated name to "Convert-RibbonSbcConfigToWord.ps1" </li>
<li>Added new bits in 8.0.0 (b502): 
<ul>
<li>NEW COLOUR SCHEME! (And added "-SonusGreen" switch for the die-hards) </li>
<li>TOD Routing </li>
<li>Load Balancing options 4 &amp; 5 to $SipLoadBalancingLookup </li>
<li>'SIP Failover Cause Codes' in SIP Sig Gps </li>
<li>'RTCP Multiplexing' in SIP Sig Gps </li>
<li>'ICE Mode' in SIP Sig Gps </li>
<li>Renamed 'Media Crypto Profiles' to 'SDES-SRTP Profiles' </li>
<li>'DTLS-SRTP Profiles' </li>
<li>'Redundancy Retry Timer' in SIP Profiles </li>
<li>Grouped Transformations, TOD, Call Routing &amp; Call Actions in new 'Call Routing' group to align with the new on-screen layout </li>
</ul>
</li>
<li>Added new test in Test-ForNull to drop a marker in the doc if an expected lookup returned $null </li>
<li>Set Word to disable spelling and grammar checking in the document. (Thanks Mike!) </li>
<li>Hid the display of "Translations" (long obsolete) </li>
<li>Fixed bugs: 
<ul>
<li>Renamed variables in the SGPeers' .split() code to make them unique. (Had accidentally reused old having copy/pasted) </li>
<li>Corrected display of 'SDES-SRTP Profiles' Derivation Rate to add the '2^' prefix for non-zero values </li>
<li>If you ran with "-do calls" it would incorrectly not show the "incomplete" warning after the TOC. ("calls" sounds like "all") </li>
</ul>
</li>
</ul>
<h4>v7.0.3 &ndash; 31st July 2018</h4>
<ul>
<li>Added new bits in 7.0.2 (b485) and SWe Lite 7.0.3 (b141): 
<ul>
<li>New SIP Sig Gp Interop Mode 'Header Transparency' </li>
<li>Changed SIP Profile's 'FQDN in Contact Header' from $EnableLookup to $SipProfileFrmHdrLookup for Teams Direct Routing </li>
</ul>
</li>
<li>Changed Media Profiles and Media Crypto Profiles to V-tables </li>
<li>Improved error reporting if the PDF exists and is open (i.e. locked) when the script goes to create it </li>
<li>Major update to the SIP Message Manipulation rules: 
<ul>
<li>Changed handling of SIP ElementDescriptors to break out the Token, Prefix and Suffix to a separate line for each </li>
<li>Added the "pre-400" messages to $SIPDescriptionLookup </li>
<li>Changed SMM rules to use $SIPDescriptionLookup when Applicable Messages = Selected Messages &amp; a SIP code is matched </li>
<li>SMM Header &amp; Request Line rules were not displaying URI &amp; URIUser parameters </li>
<li>Changed SMM rules to display the header name in Title Case </li>
<li>Added some missing descriptors to $SipElementDescElementClassLookup &amp; $SipElementDescActionLookup to improve SMM </li>
<li>Corrected display of "Ordinal" value in Header rules. Now only shows when the header name is Contact, Route, Record-Route or History-Info or PAI </li>
<li>Suppressed incorrect display of "Value" when Action = Remove </li>
</ul>
</li>
<li>Fixed bugs: 
<ul>
<li>Corrected SWeLite's Media Crypto Profiles, suppressing display of Master Key Lifetime, Lifetime Value and Derivation Rate values </li>
<li>Func Dump-ErrToScreen was not correctly dumping the error to screen when in "non-debug" mode </li>
<li>Added some missing SIP messages to the $SIPDescriptionLookup. (Thank you again Mitsu-San) </li>
<li>Updated Strip-TrailingCR to keep removing the last char from the string until the last char ISN'T a CR. (Was only stripping once). </li>
</ul>
</li>
</ul>
<h4>v7.0.1 &ndash; 14th April 2018</h4>
<ul>
<li>Added new bit in 7.0.1 (b483): 
<ul>
<li>"Encrypt AD Cache" in Auth and Directory Services / AD / Configuration </li>
</ul>
</li>
<li>Re-jigged AD / Configuration to current layout. Removed some old variables. Changed label on "AD Backup" to "AD Backup Failure Alarm" </li>
</ul>
<h3>Help me improve the script</h3>
<p>PLEASE let me know if you encounter any problems. All I ask is a copy of the .tar or symphonyconfig.xml file (de-identified as you wish), and a screen-grab from the browser to show me what it's <em>meant</em> to look like on-screen.</p>
<hr />

<br>

\- G.

<br>

This script was originally published at [https://greiginsydney.com/uxbuilder/](https://greiginsydney.com/uxbuilder/).
