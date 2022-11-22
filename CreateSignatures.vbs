
Sub MainProgram()

    DIM A, OnlyOneSignature, OneSignatureName, ForceSig	, CheckDNS 

	use_debug = true
	
	ON ERROR RESUME NEXT
		IF fileSystem.FileExists(signaturesFolderPath&"debug.txt") Then
			fileSystem.Deletefile(signaturesFolderPath&"debug.txt")
		END IF    
	on error goto 0 
	
	IF fileSystem.FileExists(signaturesFolderPath + "_Signature.ini") Then
	
		' Manually Get the Users details from the special File
		' Create the signature files in the users profile
		ManualUser = True
		AddDebug "Found _Signatuire.ini File - Read User details from this file instead of looking at ActiveDirectory"
		ReadUserInfoFile (signaturesFolderPath + "_Signature.ini")
		OneSignatureName = ""
		OnlyOneSgnature = False
		A = 1
		if len(SignatureGroupFileName(A)) > 0 then 
			CreateSignature SignatureGroupFileName(A), SignatureGroupFileName(A)+".tpl"
				
			' Force this Signature to be the Default one.
			IF ForceSignature= True then 
				ForceSig = SignatureGroupFileName(A)
			END IF
		END If
		
	ELSE
	
		ON ERROR RESUME NEXT
		Set currentUser = GetObject("LDAP://" & adSystemInfo.UserName)

		' if adSystemInfo.UserName then AddDebug "LDAP Lookup      :" + cstr(adSystemInfo.UserName)
	
		ON ERROR GOTO 0
	
		' Find out all the Signature Groups in this OU
		MaxSignatureGroup = GetSignatureGroups(LDAPurl)
		
		' Unused at the moment - Shows Black DOS Box suring users login. (Uncomment to use it)
		'CheckDNS = DNSlookup(sourceFilesUrl)
		'if CheckDNS = "" THEN 
		'	AddDebug ("DNS Lookup for " & sourceFilesUrl & " Failed - Now Using " & sourceFilesIP & " For Downloading File " & templateFileName)
		'	sourceFilesUrl = sourceFilesIP 
		'END IF
	
		' Check to see if user is in any Signature group then Create the signature for Outlook                
		IF MaxSignatureGroup > 0 THEN
			OneSignatureName = ""
			OnlyOneSgnature = False
			A = 1
			DO
				if len(SignatureGroupName(A)) > 0 then 
					IF IsMember(currentUser, SignatureGroupName(A)) Then 
				
						' Create the signature files in the users profile
						CreateSignature SignatureGroupFileName(A), SignatureGroupFileName(A)+".tpl"
				
						' Force this Signature to be the Default one.
						IF ForceSignature= True then 
							ForceSig = SignatureGroupFileName(A)
						END IF
					
						IF OnlyOneSgnature = False and len(OneSignatureName)=0 then 
							OnlyOneSgnature = True
							OneSignatureName=SignatureGroupFileName(A)
						ELSE
							' User has as more than one Signature so we cant set the default automatically.
							OnlyOneSgnature = False
						END IF
				
					ELSE
						' Remove unrequired Signature Files if they are found in the Users profile
						DeleteSignature SignatureGroupFileName(A)
					END IF
				END IF
				a=a+1
			LOOP UNTIL a> MaxSignatureGroup
			AddDebug "Signature Groups :" + cstr(MaxSignatureGroup)
		END IF

	END IF
	
	
	' A force signature setting has been set, so we will set that signature as default.
	IF len(ForceSig) <> 0 Then
		SetDefaultMailSignature(ForceSig)
		OnlyOneSgnature = False
	END If
	
	' User only has only 1 Signature - so lets set that as the default one.
	IF OnlyOneSgnature = True and len(OneSignatureName) <> 0 then 
		SetDefaultMailSignature(OneSignatureName)
	END IF

	use_debug = True
	AddDebug "Script Completed."

	SET fileSystem = Nothing
	SET shell = Nothing
	SET adSystemInfo = nothing
	SET objOU = Nothing
	SET currentUserGroups = Nothing
	
End Sub
