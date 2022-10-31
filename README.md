# Outlook Email Signature
Automatically Create Outlook Signatures on Windows Workstations based on Active Directory User Group Membership

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

Use the users Active Directory Details to manage what is displayed in the Signature, Use the info Field to Ajdust informtaion based on multiple signatures.
HTML, TEXT, VCF Outlook Signature Files will be created in the users \%appdata%\Roaming\Microsoft\Signatures\ Folder.

Any Images associated with the Signature will also be copied in to the signatures folder. If the DefaultImageType is set to png then the SignatureName.png is automatically copied. Use the AdditionalImage settings to copy other image files.

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
	-default						= Turn off Setting the Default Signature 
	-companyphone					= Removes company phone for all the users signatures
	+companyphone(XXXXX)			= Change the companyphone to XXXXXX
	-postcode						= Remove the PostCode field from the signature
	-pobox							= Remove the POBox field from the signature
	-office							= Remove the Office field from the signature
	-department						= Remove the Department field from the signature
	-state							= Remove the State field from the signature
	-suburb							= Remove the Suburb field from the signature
	-city 							= Remove the City field from the signature
	-address						= Remove the Address field from the signature
	-ipphone						= Remove the ipphone field from the signature
	-phone							= Remove the Phone field from the signature
	-mobile							= Remove the Mobile field from the signature
	-country						= Remove the Country field from the signature
	-title							= Remove the Title field from the signature
	-name							= Remove the Name field from the signature
	+name(XXXXX)					= Change the Name field to XXXXX
	-firstname						= Also removes this Value from the 'name' value if contains it.
	+firstname(XXXXX)				= Change the FirstName field to XXXXX
	-lastname						= Remove the LastName field 
	-social							= Remove the Social Icons from the signature
	+xmasmessage(XXXXX)				= Change the Xmas Message 
	-notes							= Clear the note value (use this to still keep the old data in the notes field, but it never gets used)
	testxmas						= Test The Xmas Formatting for the Signature
  
