# OutlookSignature
Automatically Create Outlook Signatures based on Active Directory User Group Membership





   DEFAULT VALUES
===================================================================================================================================================================
	Bracket Values are automatically changed like {div} {/div} = replaced with <div> </div> but only if the line does not start and end like '...'
	'.......'> Values = Remove the '' from either end. Searches for '> or ' > to determine the end of the Line

	Address		(Vbcr) chr(13) are removed 
	FooterText	^ are replaced with <p>....</p> as long as the initial value does not orginally contain any <p> HTML statements

    ACTIVE DIRECTORY
===================================================================================================================================================================

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
	

    NOTES FIELD Settings and Examples
===================================================================================================================================================================

	+signature(xxxx) [ OPTIONS ] 		= will only change OPTIONS if the signature xxxx is the same name as the current processing signature
	+signature(xxxx) +title(xxxxx)  	= will only change the title to xxxx if the signature xxxx is the same name as the current processing signature
					 +address(xxxxx) 	= will only change the address to xxxx if the signature xxxx is the same name as the current processing signature
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
  
