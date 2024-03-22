# Outlook Email Signature
Automatically Create Outlook Signatures on Windows Workstations based on Active Directory User Group Membership, Or a text file containing user details.

Place the CreateSignatures.vbs, CreateSignatures.wsf, SignatureDefaults.vbs, SignatureFunctions.vbs files into a Shared Folder EG: \\\ServerName\FileShare</br>
Place the SIGNATURE.tpl, SIGNATURE-Xmas.tpl and any Images file into a WEB Server which only needs to be accessable from the internal Network.</br>
Update the SignatureDefaults.vbs file with the Relevant ADDomain, SourceFilesURL, LDAPurl details.</br>
	ADDomain is the Base DN Path where DC = Domain Component. Use Active Directory Users and Computers. Select View, Advanced Features, Then get the properties of the first entry, and select Attribute Editor. The DN path will be shown in the distinguishedName Attribute. You can edit it and copy the value.</br>

Update the SIGNATURE.tpl file. Rename the File to reflect the Company Name. + Also update the contents of the file with the relevenat details and settings.
NOTE: The Heading Text of the File needs to include the Name of the file. If the script does not find the name of the tpl file insdie it, it wont run.</br>

NB: The HTML code is Outlook HTML 1.0 so it does not support the newer HTML commands.
Each .tpl file is split up into 2 main sections. The Signature Variables & Settings ( Set Between the {<!-- -->}) , and the Signature HTML Code.

Create Signature Group(s) and place users into these/this Group(s). Ensure that the Description Of the Group is the Name of the Signature .tpl file.
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
 *|company|*
 *|companyphone|*
 *|companyfax|*
 *|companyurl|*
 *|department|*
 *|countryname|*						Full Country Name	
 *|address|*							Users Address		(streetAddress)
 *|state|*							Users State		(st)
 *|city|*							Users city		(l)
 *|country|*							2 Digit Country Code	(c)
 *|suburb|*							Users city 		(l)
 *|postcode|*							Users postcode		(postalCode)
 *|title|*							Users Title		(title)
 *|name|*							Users Full Name		(CN)
 *|phone|*							Users phone		(telephoneNumber)
 *|mobile|*							Users Mobile		(mobile) 
 *|email|*							Users Email Address	(mail)
 *|ipphone|*							Users ip phone		(ipPhone) 
 *|pobox|*							Users pobox		(postOfficeBox)
								 ** Will add the words 'PO BOX' if this value is just a number)
 *|firstname|*							Users First Name	(givenNAme)
 *|lastname|*							Users Surname		(sn)
 *|notes|*							General Information	(info) 
								The user can have control over what data is not displayed on their signature. 
								-phone, -address, -companyphone, -mobile, -state, -country, -pobox, -city, -title
								-office, -suburb, -lastname, -firstname, -name, 
								-social = dont display social icons
								 +signature(xxxxx) +title(xxxx) +address(xxxx^xxx^xxx) = Users title changes to xxxx if Signature Name is xxxxx
 *|webpage|*							(Users WWW webpage)	(www)
 
  
 Formatting Values:						
 
 xxxxxxx_v values are simply classes that are used to hide or display values if their corresponding variables have values or not.
 EG: if the corresponsing variable is blank then the xxxxxxx_v value is set to style = 'display:none' otherwise the style ='' (nothing) 
 This means that you can use a <Span class="name_v"> <Span class="name" > ....... </span></span> to hide or show text if the name value is blank or is not.
	
 *|company_v|*							
 *|companyphone_v|*
 *|companyfax_v|*
 *|companyurl_v|*
 *|department_v|*
 *|countryname_v|*
 *|address_v|*
 *|state_v|*
 *|city_v|*
 *|country_v|*
 *|suburb_v|*								
 *|postcode_v|*
 *|title_v|*
 *|name_v|*
 *|phone_v|*
 *|mobile_v|*
 *|email_v|*
 *|ipphone_v|*
 *|pobox_v|*
 *|firstname_v|*
 *|lastname|*
 *|xmastext_v|*       
 *|IDnumber_v|*       
 *|extranote_v|* 
 *|notes_v|*  
 *|socialicons_v|* 						If the user has -social in their AD notes then this is set to " style = 'display:none' "
 
 Additional Display Values:
	
 *|day|*							Display the current Date Day Value EG 5
 *|dday|*							Display the current dates 2 digit Day Value EG 05
 *|dayname|*							Display the current dates Days Name EG Monday
 *|month|*							Display the current dates Month Value EG 1
 *|mmonth|*							Display the current dates 2 Digit Month Value EG 01
 *|monthname|*							Display the current dates Month Name EG January
 *|year|*							Display the current dates 4 Digit year value EG 2020
 *|signaturename|*						Display the Signature Name
 *|imagename|*							Set using the |ImageName| value
 *|highlightcolor|*						Set using the |HightlightColor| value
 *|height|*							Set using the |Height| value
 *|width|*							Set using the |Width| value
 *|xmastext|*							Set using the |XmasText| value
 *|IDnumber|*       						Generates a unique number based on HEX values for each Character
 *|extranote|*							Set using the Group Membership lookup with |MemberNote| value
 *|href|*							Defaults to webpage, but can get overwritten by setting DefaultHREF
 
 ** THESE VARIABLES SET VALUES **   				Propercase variables + value to use
 ** The '>' character is not allowed in any value - as this character determines the end of the value to place into the variable.
 
 These values are used if the users AD settings are not defined:
	
 <*|Company|*CompanyName>			set the company value to Company Name
 <*|CompanyPhone|*07 5555 555>			set the companyphone value to '07 5555 555'
 <*|CompanyFax|*>				set the companyfax value
 <*|CompanyURL|*>				set the companyurl value
 <*|UserPOBox|*>			 	Use this if there is no user pobox value
 <*|UserAddress|*> 				Use this if there is no user address value
 <*|UserState|*> 				Use this if there is no user state value
 <*|UserSuburb|*> 				Use this if there is no user suburb value
 <*|UserMobile|*> 				Use this if there is no user mobile value
 
 Additional Parameters:
	 
 <*|DefaultImageType|*png> 			Sets the default Image extension name to this value EG: .png
 <*|ImageName|*XXXXX.XXX> 		        Use this image (filename) in signature - If not defined then *|imagename|* variable defaults to '[signaturename].jpg'
						If an imagename value is not defined, then the script will go through all the src statements and download these files
 <*|AditionalImage|*XXXXX>			Use this Second image (filename) File in signature
 <*|HighlightColor|*#XXXXX>			Force to use this highlight colour EG: #56432
 <*|ForceEmailFirstName|*>     			Change email address to use onlt the firstname@domain EG: Yes
 <*|EmailDomain|*XXX.XXX>	  		Change email domain address to email@domain
 <*|Height|*XX>   				Force Image height EG: 150
 <*|Width|*XX>    				Force Image Width EG: 150 - must also be used in <td style="width: xx"> line to force gmail and outlook html editor compatability
 <*|ReplaceEmailDomain|*XXXX.XXX,xxxxx.xx>	Change Email Domain Name to the first 1 listed, if found in the subsequent ones
						EG: @newdomain.com.au,@previousdomain.com.au,@otherdomain.com.au
 <*|FooterText|*XXXXXXXXXXXX>			Change the FooterText Signature value to XXXXXXXXXXX
 <*|DefaultImageType|*png> 			Sets the Default Image Type to png, jpg, gif (you should not have a combination of image file types in a signature)
 
 <*|AddPhoneSpace|*True>			If True, Adds a space char to the end of ALL phone, companyphone, companyfax, ipphone, mobile, fax numbers (force it to look better)
 <*|AddAddressSpace|*True>			If True, Adds a space char to the end of ALL Address, Suburb, City, State, Country, Postcode
 <*|AddAdditionalAddressSpace|*True>		If True, Adds an additional space only Address
 <*|AddAdditionalStateSpace|*True>		If True, Adds an additional space only State
 <*|AddAdditionalCitySpace|*False>		If True, Adds an additional space only City
 <*|AddAdditionalPhoneSpaces|*True>		If True, Adds an additional space only Phone
 
 <*|AutoPOBox|*True> 				If True, Modifyes the POBOX field to include the Words 'PO Box' if pobox value is just a number.
 <*|ForceCreate|*True>				If True, Force this signature to be updated everytime the user logges in.											
 <*|ForceDefaultSignature|*True>		If True, Force this siganture to be the default.
						If the user only has 1 signature then that signature is set as default.
	
  <*|MemberNote|*LDAP://CN=SG-OutlookSignatureNOTE1,OU=Signatures,OU=Groups,OU=AD - Homecorp Constructions,DC=ad,DC=homecorp,DC=com>  		
  
 Set Social Icons Values:													
 <*|SocialIconName|*ImageName>				Modifies the Socialicon Image Name to use this name instead.
 <*|SocialIcon|*FaceBook>				Modifies the *|FaceBook|* value with the TemplateName-FaceBook.png Image.
 <*|SocialIconLink|*https://www.facebook.com/> 		*|FaceBookLink|*
 <*|SocialIconAlt|*Homecorp Facebook>   		Adds to the Text  *|FaceBookAlt|* 
 <*|SocialIcon|*Instagram>
 <*|SocialIconLink|*https://www.instagram.com>
 <*|SocialIcon|*LinkedIn>
 <*|SocialIconLink|*https://www.linkedin.com>
 <*|SocialIcon|*YouTube>				*|YouTube|*   = TemplateName-YouTube.png
 <*|SocialIconLink|*https://www.youtube.com>		*|YouTubeLink|*
 
 Add HTML Code to Field Values:
 <*|InsertHTML|*address>						Changes the value of address, notes,  fields
 <*|InsertHTMLFind|*char(10)>						Finds this text in the field - use 'char(x)' to find  the charater ascii value
 <*|InsertHTMLCode|*'<span class="Bar">|</span>'> 			to this value if found in the field. 
 
 This is the xmas option settings:
 <*|XmasFrom|*1/12/2019>         					dd/mm/yyyy)   or (1/12/YYYY - If jan then Sets XmasFrom Xmasto Year -1 
 <*|XmasTo|*12/1/2020>			 				dd/mm/yyyy)   or (12/1/YYYY - set Year to the XmaxsFrom value +1
 <*|XmasImageName|*filename.jpg>					Xmas ImageFile to use - if not defined the uses 'SignatureName-Xmas.jpg'
 <*|XmasText|*Our office will be closed from 4.00pm on Friday 20th December 2019. We will be reopen at 8:30am on Monday 13th January 2020 with skeleton staff operating from Monday 6th January 2020.>
 <*|XmasHighlightColor|*#0075C9>					Changes |highlightcolor| to be this if between the christmas from to period)
 <*|XmasOverWriteColor|*TextColor,LinkWWWColor>				Changes all the occurances of TextHighlightColor or LinkWWWColor with the XmasHightlightColor value
 <*|MemberNote|*LDAP://CN=SG-OutlookSignatureNOTE1,OU=OutlookSignatureNotes,OU=Signatures,OU=Groups,OU=DOMAIN,DC=ad,DC=domainname,DC=com,DC=au> If the user is a member if this Group then the |extranote| = Group Description Text
											
 Force Users AD values to these values:				(Overwrites the users current values)
 <*|DefaultCompany|*xxxxxxxx>					Force Company to use this value
 <*|DefaultPhone|*xxxxxxx>					Force CompanyPhone to this value	
 <*|DefaultFax|*>						Force CompanyFax to this value
 <*|DefaultEmailDomain|*>					Force users email domain to now use this value
 <*|DefaultPOBox|*>						Force pobox to this value
 <*|DefaultTitle|*"">						Force users title to be this value. If "" then blanks the value
 <*|DefaultName|*"">						Force users name to be this value. If "" then blanks the value
 <*|DefaultEmail|*"">						Force users email address to be this value. If "" then blanks the value
 <*|DefaultState|*"">						Force users state to be this value. If "" then blanks the value
 <*|DefaultAddress|*"">						Force users address to be this value. If "" then blanks the value
 <*|DefaultMobile|*"">						Force users mobile to be this value. If "" then blanks the value
 <*|DefaultSuburb|*"">						Force users suburb to be this value. If "" then blanks the value
 <*|DefaultWWW|*xxxx>						Force user webpage URL to be xxxx. If "" then blanks the value
 <*|DefaultHREF*xxxx>						Forces the *|href|* value to be xxxx

								Email
								EmailDomain
													
 <*|InternationalPrefix|*+61>					Sets the internation prefix to use, if transforming numbers using INTPHONE, INTMOBILE
 
 <*|TransformName|*PROPER>					Converts the name to propercase-(PROPPER, UPPER, LOWER, MOBILE, PHONE, FULLNAME)
 <*|TransformTitle|*UPPER>					Converts the title to Uppercase
 <*|TransformEmail|*LOWER>					Converts the email to Lowercase
 <*|TransformState|*UPPER>					Converts the state to Uppercase
 <*|TransformCity|*PROPER>					Converts the city to Propercase
 <*|TransformMobile|*MOBILE>					Converts the number to 04XX XXX XXXX
 <*|TransformPhone|*PHONE>					Converts the name to 07 XXXX XXXX
 <*|TransformMobile|*INTMOBILE>					Converts the number to +61 4XX XXX XXXX
 <*|TransformPhone|*INTPHONE>					Converts the name to +61 7 XXXX XXXX
 <*|TransformState|*FULLNAME>					Converts the state form VIC to Victoria
 <*|TransformState|*SHORTNAME>					Converts the state form Victoria to VIC
 <*|TransformCountry|*FULLNAME>					Converts the country form AU to Australia
 <*|TransformCountry|*SHORTNAME>				Converts the country form Australia to AU
  
 <*|HideTableRows|*True>					Forces the use of the <tr class="xtranotes_t","xmastext_t" to be hidden or not - if xtranotes_v is display:none 
 
 Values that can be set within the Signature HTML or CSS:
 <*|LinkWWWColor|*#C2D500>
 <*|SymbolColor|*#002856>
 <*|TextColor|*#000000>
 <*|TextColorHighlight|*#002856>
 <*|TextFooterColor|*#666666>
 <*|BarColor|*#C2D500>
 <*|HyperLinkColor|*#0563C1>
 <*|DefaultFont|*"Calibri Light", sans-serif>
 <*|DefaultFontSize|*11.0pt>
 
  
 Combine some fields together and also use HTML code to configure what the combination looks like:
 <*|Combine|*address+state+postcode>				These are the fields that will be combined together. The result is placed into the first field
 <*|CombineIfBlank|*False> 							Add the Field if its a blank value
 <*|CombineLastCode|*False>							Combine the last fields Code or not. (just add the fields value, if not defined)
 <*|CombineField|*address>							This is the field name that is going to be used.
 <*|CombineHTMLCode|*''>							This code will be wrapped around the field when it is added
 <*|CombineField|*+state>							The next field to add. + code added after field, - code added before, # code and field merged
 <*|CombineHTMLCode|*'<span class="bar">|</span>&nbsp;&nbsp;'>	
 <*|CombineField|*+postcode>
 <*|CombineHTMLCode|*'<span class="bar">|</span>&nbsp;&nbsp;'>
 
 <*|Combine|*companyphone+phone+mobile>
 <*|CombineLastCode|*True>
 <*|CombineField|*#companyphone>
 <*|CombineHTMLCode|*'<span class="highlight">P&nbsp;</span><span class="phone"><a href="tel:*|companyphone|*"><span style="color:#212121;text-decoration:none;text-underline:none">*|companyphone|*</span></a></span>&nbsp;&nbsp;'>
 <*|CombineField|*#phone>
 <*|CombineHTMLCode|*'<span class="highlight">D&nbsp;</span><span class="phone"><a href="tel:*|phone|*"><span style="color:#212121;text-decoration:none;text-underline:none">*|phone|*</span></a></span>&nbsp;&nbsp;'>
 <*|CombineField|*#mobile>
 <*|CombineHTMLCode|*'<span class="highlight">M&nbsp;</span><span class="phone"><a href="tel:*|mobile|*"><span style="color:#212121; text-decoration:none;text-underline:none">*|mobile|*</span></a></span>&nbsp;&nbsp;'>
 
mso x - Outlook Version Numbers:
Outlook 2000 - Version 9
Outlook 2002 - Version 10
Outlook 2003 - Version 11
Outlook 2007 - Version 12
Outlook 2010 - Version 14
Outlook 2013 - Version 15
Outlook 2016 - Version 16

**HTML Coding Examples**
Remove hovering mouse with cursor:default !important
Remove Hyperlink underline with href="" + text-decoration:none !important
Change Underline color with style
<a href="" style="color:#ccc;text-decoration:none !important;cursor:default !important">www.privium.com.au</a>
<a href="" style="text-decoration: none;">

Tables:
Use Height in table TR and TD to align table row height
<tr style='height:28px'> 
<td width="100%" style='width:100.0%;padding:0cm 0cm 0cm 0cm;height:28px'>

Outlook 2016 Ignores Margin and Padding on Images
Text Doesn’t Wrap Automatically in Outlook 2016


Set Debug=True in the tpl file to create a 'debug.txt' file that information on what settings were set and modified in each signature.

**DEFAULT VALUES**

	Bracket Values are automatically changed like {div} {/div} = replaced with <div> </div> but only if the line does not start and end like '...'
	'.......'> Values = Remove the '' from either end. Searches for '> or ' > to determine the end of the Line.
	Address		(Vbcr) chr(13) are removed from the value.
	FooterText	^ are replaced with <p>....</p> as long as the initial value does not orginally contain any <p> HTML statements.

**ACTIVE DIRECTORY**

	       name = displayName
	      title = title
	      phone = telephoneNumber
	     mobile = mobile
	      email = mail
	    address = streetAddress
	      pobox = postOfficeBox
	      state = st
	       city = l
   	     suburb = l
	    country = c
	   postcode = postalCode
	     office = physicalDeliveryOfficeName
	    webpage = wWWHomePage
	countryname = co
	 department = department
      firstname = givenName
	   lastname = sn
	    ipphone = ipPhone
	WhenChanged = WhenChanged
	      notes = info
	
NOTES Field Settings and Examples.

	+signature(xxxx) [ OPTIONS ] 		= will only change OPTIONS if the signature xxxx is the same name as the current processing signature
	+signature(xxxx) +title(xxxxx)  	= will only change the title to xxxx if the signature xxxx is the same name as the current processing signature
					 +address(xxxxx) 	= will only change the address to xxxx if the signature xxxx is the same name as the current processing 								signature
					 +state(xxxx)		
					 +postcode(xxxxx)

	+xmasmessage(xxxxx) +signature(xxxxx) 
	
	+address(xxxx^xxx^xxx) 			= Change the address to XXXX - ^ will be changed to char(13)/(Carrage Return)
	-default				= Turn off Setting the Default Signature 
	-companyphone				= Removes company phone for all the users signatures
	+companyphone(XXXXX)			= Change the companyphone to XXXXXX
	-postcode				= Remove the PostCode field from the signature
	-pobox					= Remove the POBox field from the signature
	-office					= Remove the Office field from the signature
	-department				= Remove the Department field from the signature
	-state					= Remove the State field from the signature
	-suburb					= Remove the Suburb field from the signature
	-city 					= Remove the City field from the signature
	-address				= Remove the Address field from the signature
	-ipphone				= Remove the ipphone field from the signature
	-phone					= Remove the Phone field from the signature
	-mobile					= Remove the Mobile field from the signature
	-country				= Remove the Country field from the signature
	-title					= Remove the Title field from the signature
	-name					= Remove the Name field from the signature
	+name(XXXXX)				= Change the Name field to XXXXX
	-firstname				= Also removes this Value from the 'name' value if contains it.
	+firstname(XXXXX)			= Change the FirstName field to XXXXX
	-lastname				= Remove the LastName field 
	-social					= Remove the Social Icons from the signature
	+xmasmessage(XXXXX)			= Change the Xmas Message 
	-notes					= Clear the note value (use this to still keep the old data in the notes field, but it never gets used)
	testxmas				= Test The Xmas Formatting for the Signature
  
