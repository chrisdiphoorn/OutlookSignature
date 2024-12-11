# Outlook Email Signature
Automatically Create Outlook Signatures based on Active Directory User Group Membership, Or a text file containing user details.

Place the CreateSignatures.vbs, CreateSignatures.wsf, SignatureDefaults.vbs, SignatureFunctions.vbs files into a Shared Folder EG: \\\ServerName\FileShare</br>
Place the SIGNATURE.tpl, SIGNATURE-Xmas.tpl and any Image files into a WEB Server. Create a vhost site that only needs to be accessable from the internal Network.</br>
Update the SignatureDefaults.vbs file with the Relevant ADDomain, SourceFilesURL, LDAPurl details.</br>
	ADDomain is the Base DN Path where DC = Domain Component. Use Active Directory Users and Computers. Select View, Advanced Features, Then get the properties of the first entry, and select Attribute Editor. The DN path will be shown in the distinguishedName Attribute. You can edit it and copy the value.</br>

Update the SIGNATURE.tpl file. Rename the File to reflect the Company Name and also update the contents of the file with the relevenat details and settings.
NOTE: The Heading Text of the File needs to include the Name of the file. If the script does not find the name of the tpl file insdie it, it wont run.</br>

NB: The HTML code is based on Outlook HTML 1.0 so it does not support any newer HTML commands. Outlook HTML can be a little tight and finiky on how to get things 100%.
A xxxxxx.tpl file is split up into 2 main sections. The Signature Variables & Settings ( Set Between the {<!-- -->}) , and the Signature HTML Code.

In Activdirectory create Signature Group(s) and place users into these/this Group(s). Ensure that the Description Of the Group is the Name of the Signature .tpl file.
Update the OU details into the SignatureDeafult file so the script knows where to look for the Signature Group(s).

From the Users Loginscript run the c:\windows\system32\cscript.exe //NoLogo \\ServerName\FileShare\CreateSignatures.wsf File.

Use the users Active Directory Details to manage what is displayed in the Signature. EG Title, Mobile, Name, Address ....
You can use the Users AD "info" Field to Ajdust information based on multiple signatures. So the user can have different title / address / name when using more than one signature.

The script will produce Outlook compatable files SignatureName.HTML, SignatureName.TEXT, SignatureName.VCF Outlook Signature Files will be created in the users \%appdata%\Roaming\Microsoft\Signatures\ Folder.

Any Images associated with the Signature will also be copied in to the signatures folder. for the Signature = XXXX.tpl, any image files with same name as the signature EG XXXX.jpg, XXXX.png, will be automatically copied to the users signature folder. You may set the 'DefaultImageType' is set to png then the XXXX.png is automatically copied. 
Use the 'AdditionalImage' setting to ensure that these images are also copied to the users Oulook Signature Folder copy other image files.

**TPL FILE ATTRIBUTES**

To use each parameter add a '<' at the beginning so the parameter looks like this <*|parameter|*
 Then add the Value Text, and then a '>' to show the end of the value. These settings are always set in the remarked out HTML code as per below.
 They must be used with the exact format / caseing as per below.
 Active directory fields must have a '*|' + fieldname + '|*' as per below.
 
 FIRST:
 Use an online Editor like https://html-online.com/editor/ to create a HTML format to start with?
 
 Use this link to ensure that the HTML formatting code used in the signature conforms to office 2007 rendering capabilities. (EG: Does not support HTML5)
 https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa338201(v=office.12)?redirectedfrom=MSDN
 
 THEN:
 NB*** If the Users Active Directory Last Update (Date + Time) has not changed then a signature will not be (re)created. Unless the ForceCreate Option is True
 NB*** If the user is not a member of a Signature Group then a Signature will not be created.
 NB*** If signature files are found and the user is not a member of the group then the files will be deleted.
 Add These template files to the beginning of the HTML document, Rename the document to .tpl and modify with a Text Editor.
 The HTML lines below are just for reference only and can be used as a simple reference to how to use it.
 
  You may use the values of '-mobile' or '-ipphone' in ActiveDirectory users Notes Field to ensure that their mobile number does not show in all their signatures
 Or +signature(SignatureName) +title(New Title) to update the title for the Specific Signatrure
 Or +signature(SignatureName) +address(New Address) +postcode(New Postcode) +pobox( New PoBox) for the Specific Signature
 The engine will change the values before they get used in the signature.
 
 ** THESE VARIABLES DISPLAY DATA **  (lowercase variables)
 
 Active Directory Field Values:
 *|company|*</br>
 *|companyphone|*</br>
 *|companyfax|*</br>
 *|companyurl|*</br>
 *|department|*</br>
 *|countryname|*						Full Country Name	</br>
 *|address|*							Users Address		(streetAddress)</br>
 *|state|*							Users State		(st)</br>
 *|city|*							Users city		(l)</br>
 *|country|*							2 Digit Country Code	(c)</br>
 *|suburb|*							Users city 		(l)</br>
 *|postcode|*							Users postcode		(postalCode)</br>
 *|title|*							Users Title		(title)</br>
 *|name|*							Users Full Name		(CN)</br>
 *|phone|*							Users phone		(telephoneNumber)</br>
 *|mobile|*							Users Mobile		(mobile) </br>
 *|email|*							Users Email Address	(mail) </br>
 *|ipphone|*							Users ip phone		(ipPhone) </br>
 *|pobox|*							Users pobox		(postOfficeBox) </br>
								 ** Will add the words 'PO BOX' if this value is just a number) </br>
 *|firstname|*							Users First Name	(givenName) </br>
 *|lastname|*							Users Surname		(sn) </br>
 *|notes|*							General Information	(info) </br>
								The user can have control over what data is not displayed on their signature. </br>
								-phone, -address, -companyphone, -mobile, -state, -country, -pobox, -city, -title </br>
								-office, -suburb, -lastname, -firstname, -name, </br>
								-social = dont display social icons </br>
								 +signature(xxxxx) +title(xxxx) +address(xxxx^xxx^xxx) = Users title changes to xxxx if Signature Name is xxxxx </br>
 *|webpage|*							(Users WWW webpage)	(www) </br>
 </br>
  
 Formatting Values:						</br>
 </br>
 xxxxxxx_v values are simply classes that are used to hide or display values if their corresponding variables have values or not. </br>
 EG: if the corresponsing variable is blank then the xxxxxxx_v value is set to style = 'display:none' otherwise the style ='' (nothing) </br> 
 This means that you can use a <Span class="name_v"> <Span class="name" > ....... </span></span> to hide or show text if the name value is blank or is not. </br>
	
 *|company_v|*		</br>
 *|companyphone_v|*	</br>
 *|companyfax_v|*	</br>
 *|companyurl_v|*	</br>
 *|department_v|*	</br>
 *|countryname_v|*	</br>
 *|address_v|*		</br>
 *|state_v|*		</br>
 *|city_v|*		</br>
 *|country_v|*		</br>
 *|suburb_v|*		</br>						
 *|postcode_v|*		</br>
 *|title_v|*		</br>
 *|name_v|*		</br>
 *|phone_v|*		</br>
 *|mobile_v|*		</br>
 *|email_v|*		</br>
 *|ipphone_v|*		</br>
 *|pobox_v|*		</br>
 *|firstname_v|*	</br>
 *|lastname|*		</br>
 *|xmastext_v|*       	</br>
 *|IDnumber_v|*       	</br>
 *|extranote_v|* 	</br>
 *|notes_v|*  		</br>
 *|socialicons_v|* 	If the user has -social in their AD notes then this is set to " style = 'display:none' "</br>
 </br>
 Additional Display Values:</br>
	</br>
 *|day|*							Display the current Date Day Value EG 5</br>
 *|dday|*							Display the current dates 2 digit Day Value EG 05</br>
 *|dayname|*							Display the current dates Days Name EG Monday</br>
 *|month|*							Display the current dates Month Value EG 1</br>
 *|mmonth|*							Display the current dates 2 Digit Month Value EG 01</br>
 *|monthname|*							Display the current dates Month Name EG January</br>
 *|year|*							Display the current dates 4 Digit year value EG 2020</br>
 *|signaturename|*						Display the Signature Name</br>
 *|imagename|*							Set using the |ImageName| value</br>
 *|highlightcolor|*						Set using the |HightlightColor| value</br>
 *|height|*							Set using the |Height| value</br>
 *|width|*							Set using the |Width| value</br>
 *|xmastext|*							Set using the |XmasText| value</br>
 *|IDnumber|*       						Generates a unique number based on HEX values for each Character</br>
 *|extranote|*							Set using the Group Membership lookup with |MemberNote| value</br>
 *|href|*							Defaults to webpage, but can get overwritten by setting DefaultHREF</br>
 </br>
 ** THESE VARIABLES SET VALUES **   				Propercase variables + value to use</br>
 ** The '>' character is not allowed in any value - as this character determines the end of the value to place into the variable.</br>
 </br>
 These values are used if the users AD settings are not defined:</br>
	
 <*|Company|*CompanyName>			set the company value to Company Name </br>
 <*|CompanyPhone|*07 5555 555>			set the companyphone value to '07 5555 555' </br>
 <*|CompanyFax|*>				set the companyfax value </br>
 <*|CompanyURL|*>				set the companyurl value </br>
 <*|UserPOBox|*>			 	Use this if there is no user pobox value </br>
 <*|UserAddress|*> 				Use this if there is no user address value </br>
 <*|UserState|*> 				Use this if there is no user state value </br>
 <*|UserSuburb|*> 				Use this if there is no user suburb value </br>
 <*|UserMobile|*> 				Use this if there is no user mobile value </br>
 </br>
 Additional Parameters:</br>
	 </br>
 <*|DefaultImageType|*png> 			Sets the default Image extension name to this value EG: .png </br>
 <*|ImageName|*XXXXX.XXX> 		        Use this image (filename) in signature - If not defined then *|imagename|* variable defaults to '[signaturename].jpg' </br>
						If an imagename value is not defined, then the script will go through all the src statements and download these files </br>
 <*|AditionalImage|*XXXXX>			Use this Second image (filename) File in signature </br>
 <*|HighlightColor|*#XXXXX>			Force to use this highlight colour EG: #56432 </br>
 <*|ForceEmailFirstName|*>     			Change email address to use onlt the firstname@domain EG: Yes </br>
 <*|EmailDomain|*XXX.XXX>	  		Change email domain address to email@domain </br>
 <*|Height|*XX>   				Force Image height EG: 150 </br>
 <*|Width|*XX>    				Force Image Width EG: 150 - must also be used in <td style="width: xx"> line to force gmail and outlook html editor compatability </br>
 <*|ReplaceEmailDomain|*XXXX.XXX,xxxxx.xx>	Change Email Domain Name to the first 1 listed, if found in the subsequent ones </br>
						EG: @newdomain.com.au,@previousdomain.com.au,@otherdomain.com.au </br>
 <*|FooterText|*XXXXXXXXXXXX>			Change the FooterText Signature value to XXXXXXXXXXX </br>
 <*|DefaultImageType|*png> 			Sets the Default Image Type to png, jpg, gif (you should not have a combination of image file types in a signature) </br>
 </br>
 <*|AddPhoneSpace|*True>			If True, Adds a space char to the end of ALL phone, companyphone, companyfax, ipphone, mobile, fax numbers (force it to look better) </br>
 <*|AddAddressSpace|*True>			If True, Adds a space char to the end of ALL Address, Suburb, City, State, Country, Postcode </br>
 <*|AddAdditionalAddressSpace|*True>		If True, Adds an additional space only Address </br>
 <*|AddAdditionalStateSpace|*True>		If True, Adds an additional space only State </br>
 <*|AddAdditionalCitySpace|*False>		If True, Adds an additional space only City </br>
 <*|AddAdditionalPhoneSpaces|*True>		If True, Adds an additional space only Phone </br>
 </br>
 <*|AutoPOBox|*True> 				If True, Modifyes the POBOX field to include the Words 'PO Box' if pobox value is just a number. </br>
 <*|ForceCreate|*True>				If True, Force this signature to be updated everytime the user logges in.</br>											
 <*|ForceDefaultSignature|*True>		If True, Force this siganture to be the default.</br>
						If the user only has 1 signature then that signature is set as default.</br>
	</br>
  <*|MemberNote|*LDAP://CN=SG-OutlookSignatureNOTE1,OU=Signatures,OU=Groups,OU=AD - Homecorp Constructions,DC=ad,DC=homecorp,DC=com>  		</br>
  </br>
 Set Social Icons Values:													</br>
 <*|SocialIconName|*ImageName>				Modifies the Socialicon Image Name to use this name instead. </br>
 <*|SocialIcon|*FaceBook>				Modifies the *|FaceBook|* value with the TemplateName-FaceBook.png Image. </br>
 <*|SocialIconLink|*https://www.facebook.com/> 		*|FaceBookLink|* </br>
 <*|SocialIconAlt|*Homecorp Facebook>   		Adds to the Text  *|FaceBookAlt|*  </br>
 <*|SocialIcon|*Instagram> </br>
 <*|SocialIconLink|*https://www.instagram.com></br>
 <*|SocialIcon|*LinkedIn></br>
 <*|SocialIconLink|*https://www.linkedin.com></br>
 <*|SocialIcon|*YouTube>				*|YouTube|*   = TemplateName-YouTube.png</br>
 <*|SocialIconLink|*https://www.youtube.com>		*|YouTubeLink|*</br>
 </br>
 Add HTML Code to Field Values:</br>
 <*|InsertHTML|*address>						Changes the value of address, notes,  fields</br>
 <*|InsertHTMLFind|*char(10)>						Finds this text in the field - use 'char(x)' to find  the charater ascii value</br>
 <*|InsertHTMLCode|*'<span class="Bar">|</span>'> 			to this value if found in the field. </br>
 </br>
 This is the xmas option settings:</br>
 <*|XmasFrom|*1/12/2019>         					dd/mm/yyyy)   or (1/12/YYYY - If jan then Sets XmasFrom Xmasto Year -1 </br>
 <*|XmasTo|*12/1/2020>			 				dd/mm/yyyy)   or (12/1/YYYY - set Year to the XmaxsFrom value +1 </br>
 <*|XmasImageName|*filename.jpg>					Xmas ImageFile to use - if not defined the uses 'SignatureName-Xmas.jpg' </br>
 <*|XmasText|*Our office will be closed from 4.00pm on Friday 20th December 2019. We will be reopen at 8:30am on Monday 13th January 2020 with skeleton staff operating from Monday 6th January 2020.> </br>
 <*|XmasHighlightColor|*#0075C9>					Changes |highlightcolor| to be this if between the christmas from to period) </br>
 <*|XmasOverWriteColor|*TextColor,LinkWWWColor>				Changes all the occurances of TextHighlightColor or LinkWWWColor with the XmasHightlightColor value </br>
 <*|MemberNote|*LDAP://CN=SG-OutlookSignatureNOTE1,OU=OutlookSignatureNotes,OU=Signatures,OU=Groups,OU=DOMAIN,DC=ad,DC=domainname,DC=com,DC=au> If the user is a member if this Group then the |extranote| = Group Description Text </br>
											</br>
 Force Users AD values to these values:				(Overwrites the users current values)</br>
 <*|DefaultCompany|*xxxxxxxx>					Force Company to use this value </br>
 <*|DefaultPhone|*xxxxxxx>					Force CompanyPhone to this value </br>
 <*|DefaultFax|*>						Force CompanyFax to this value </br>
 <*|DefaultEmailDomain|*>					Force users email domain to now use this value </br>
 <*|DefaultPOBox|*>						Force pobox to this value </br>
 <*|DefaultTitle|*"">						Force users title to be this value. If "" then blanks the value </br>
 <*|DefaultName|*"">						Force users name to be this value. If "" then blanks the value </br>
 <*|DefaultEmail|*"">						Force users email address to be this value. If "" then blanks the value </br>
 <*|DefaultState|*"">						Force users state to be this value. If "" then blanks the value </br>
 <*|DefaultAddress|*"">						Force users address to be this value. If "" then blanks the value </br>
 <*|DefaultMobile|*"">						Force users mobile to be this value. If "" then blanks the value </br>
 <*|DefaultSuburb|*"">						Force users suburb to be this value. If "" then blanks the value </br>
 <*|DefaultWWW|*xxxx>						Force user webpage URL to be xxxx. If "" then blanks the value </br>
 <*|DefaultHREF*xxxx>						Forces the *|href|* value to be xxxx </br>
</br>
								Email </br>
								EmailDomain </br>
													
 <*|InternationalPrefix|*+61>					Sets the internation prefix to use, if transforming numbers using INTPHONE, INTMOBILE </br>
 </br>
 <*|TransformName|*PROPER>					Converts the name to propercase-(PROPPER, UPPER, LOWER, MOBILE, PHONE, FULLNAME) </br>
 <*|TransformTitle|*UPPER>					Converts the title to Uppercase </br>
 <*|TransformEmail|*LOWER>					Converts the email to Lowercase </br>
 <*|TransformState|*UPPER>					Converts the state to Uppercase </br>
 <*|TransformCity|*PROPER>					Converts the city to Propercase </br>
 <*|TransformMobile|*MOBILE>					Converts the number to 04XX XXX XXXX </br>
 <*|TransformPhone|*PHONE>					Converts the name to 07 XXXX XXXX </br>
 <*|TransformMobile|*INTMOBILE>					Converts the number to +61 4XX XXX XXXX </br>
 <*|TransformPhone|*INTPHONE>					Converts the name to +61 7 XXXX XXXX </br>
 <*|TransformState|*FULLNAME>					Converts the state form VIC to Victoria </br>
 <*|TransformState|*SHORTNAME>					Converts the state form Victoria to VIC </br>
 <*|TransformCountry|*FULLNAME>					Converts the country form AU to Australia </br>
 <*|TransformCountry|*SHORTNAME>				Converts the country form Australia to AU </br>
   </br>
 <*|HideTableRows|*True>					Forces the use of the <tr class="xtranotes_t","xmastext_t" to be hidden or not - if xtranotes_v is display:none  </br>
 </br>
 Values that can be set within the Signature HTML or CSS: </br>
 <*|LinkWWWColor|*#C2D500> </br>
 <*|SymbolColor|*#002856> </br>
 <*|TextColor|*#000000> </br>
 <*|TextColorHighlight|*#002856> </br>
 <*|TextFooterColor|*#666666> </br>
 <*|BarColor|*#C2D500> </br>
 <*|HyperLinkColor|*#0563C1> </br>
 <*|DefaultFont|*"Calibri Light", sans-serif> </br>
 <*|DefaultFontSize|*11.0pt> </br>
 </br>
  
 Combine some fields together and also use HTML code to configure what the combination looks like: </br>
 <*|Combine|*address+state+postcode>				These are the fields that will be combined together. The result is placed into the first field </br>
 <*|CombineIfBlank|*False> 							Add the Field if its a blank value </br>
 <*|CombineLastCode|*False>							Combine the last fields Code or not. (just add the fields value, if not defined) </br>
 <*|CombineField|*address>							This is the field name that is going to be used. </br>
 <*|CombineHTMLCode|*''>							This code will be wrapped around the field when it is added </br>
 <*|CombineField|*+state>							The next field to add. + code added after field, - code added before, # code and field merged </br>
 <*|CombineHTMLCode|*'<span class="bar">|</span>&nbsp;&nbsp;'>	</br>
 <*|CombineField|*+postcode> </br>
 <*|CombineHTMLCode|*'<span class="bar">|</span>&nbsp;&nbsp;'> </br>
 </br>
 <*|Combine|*companyphone+phone+mobile> </br>
 <*|CombineLastCode|*True> </br>
 <*|CombineField|*#companyphone> </br>
 <*|CombineHTMLCode|*'<span class="highlight">P&nbsp;</span><span class="phone"><a href="tel:*|companyphone|*"><span style="color:#212121;text-decoration:none;text-underline:none">*|companyphone|*</span></a></span>&nbsp;&nbsp;'> </br>
 <*|CombineField|*#phone> </br>
 <*|CombineHTMLCode|*'<span class="highlight">D&nbsp;</span><span class="phone"><a href="tel:*|phone|*"><span style="color:#212121;text-decoration:none;text-underline:none">*|phone|*</span></a></span>&nbsp;&nbsp;'></br> 
 <*|CombineField|*#mobile> </br>
 <*|CombineHTMLCode|*'<span class="highlight">M&nbsp;</span><span class="phone"><a href="tel:*|mobile|*"><span style="color:#212121; text-decoration:none;text-underline:none">*|mobile|*</span></a></span>&nbsp;&nbsp;'> </br>
 </br>
mso x - Outlook Version Numbers: </br>
Outlook 2000 - Version 9 </br>
Outlook 2002 - Version 10 </br>
Outlook 2003 - Version 11 </br>
Outlook 2007 - Version 12 </br>
Outlook 2010 - Version 14 </br>
Outlook 2013 - Version 15 </br>
Outlook 2016 - Version 16 </br>
</br>
**HTML Coding Examples** </br>
Remove hovering mouse with cursor:default !important </br>
Remove Hyperlink underline with href="" + text-decoration:none !important </br>
Change Underline color with style </br>
<a href="" style="color:#ccc;text-decoration:none !important;cursor:default !important">www.domain.com.au</a> </br>
<a href="" style="text-decoration: none;"> </br>
</br>
Tables: </br>
Use Height in table TR and TD to align table row height </br>
<tr style='height:28px'>  </br>
<td width="100%" style='width:100.0%;padding:0cm 0cm 0cm 0cm;height:28px'> </br>
</br>
Outlook 2016 Ignores Margin and Padding on Images </br>
Text Doesn’t Wrap Automatically in Outlook 2016 </br>
</br>
Set Debug=True in the tpl file to create a 'debug.txt' file that information on what settings were set and modified in each signature. </br>

**DEFAULT VALUES**</br>

	Bracket Values are automatically changed like {div} {/div} = replaced with <div> </div> but only if the line does not start and end like '...' </br>
	'.......'> Values = Remove the '' from either end. Searches for '> or ' > to determine the end of the Line. </br>
	Address		(Vbcr) chr(13) are removed from the value. </br>
	FooterText	^ are replaced with <p>....</p> as long as the initial value does not orginally contain any <p> HTML statements. </br>

**ACTIVE DIRECTORY**</br>

	       name = displayName </br>
	      title = title </br>
	      phone = telephoneNumber </br>
	     mobile = mobile </br>
	      email = mail </br>
	    address = streetAddress </br>
	      pobox = postOfficeBox </br>
	      state = st </br>
	       city = l </br>
   	     suburb = l </br>
	    country = c </br>
	   postcode = postalCode </br>
	     office = physicalDeliveryOfficeName </br>
	    webpage = wWWHomePage </br>
	countryname = co </br>
	 department = department </br>
      firstname = givenName </br>
	   lastname = sn </br>
	    ipphone = ipPhone </br>
	WhenChanged = WhenChanged </br>
	      notes = info </br>
	</br>
NOTES Field Settings and Examples.</br>
</br>
	+signature(xxxx) [ OPTIONS ] 		= will only change OPTIONS if the signature xxxx is the same name as the current processing signature </br>
	+signature(xxxx) +title(xxxxx)  	= will only change the title to xxxx if the signature xxxx is the same name as the current processing signature </br>
	+address(xxxxx) 	= will only change the address to xxxx if the signature xxxx is the same name as the current processing signature</br>
	+state(xxxx)		</br>
	+postcode(xxxxx)	</br>
	+xmasmessage(xxxxx) 	</br>
 	+signature(xxxxx)  	</br>
	+address(xxxx^xxx^xxx) 			= Change the address to XXXX - ^ will be changed to char(13)/(Carrage Return) </br>
	-default				= Turn off Setting the Default Signature  </br>
	-companyphone				= Removes company phone for all the users signatures </br>
	+companyphone(XXXXX)			= Change the companyphone to XXXXXX </br>
	-postcode				= Remove the PostCode field from the signature </br>
	-pobox					= Remove the POBox field from the signature </br>
	-office					= Remove the Office field from the signature </br> 
	-department				= Remove the Department field from the signature </br>
	-state					= Remove the State field from the signature </br>
	-suburb					= Remove the Suburb field from the signature </br>
	-city 					= Remove the City field from the signature </br>
	-address				= Remove the Address field from the signature </br>
	-ipphone				= Remove the ipphone field from the signature </br>
	-phone					= Remove the Phone field from the signature </br>
	-mobile					= Remove the Mobile field from the signature </br>
	-country				= Remove the Country field from the signature </br>
	-title					= Remove the Title field from the signature </br>
	-name					= Remove the Name field from the signature </br>
	+name(XXXXX)				= Change the Name field to XXXXX </br>
	-firstname				= Also removes this Value from the 'name' value if contains it. </br>
	+firstname(XXXXX)			= Change the FirstName field to XXXXX </br>
	-lastname				= Remove the LastName field </br>
	-social					= Remove the Social Icons from the signature </br>
	+xmasmessage(XXXXX)			= Change the Xmas Message </br>
	-notes					= Clear the note value (use this to still keep the old data in the notes field, but it never gets used) </br>
	testxmas				= Test The Xmas Formatting for the Signature </br>
  
