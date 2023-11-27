' ============================================================================================
' Author:       Chris Diphoorn 
' Create date:  23/07/2020
' Description:  Sets the Domain Signature Varibles
' Engine:		Joiitech Signature Engine	
' ============================================================================================
'
' LinkImages = True : Modifies all the image src locations to include the full Source URL instead of downloading each file.
'
Option Explicit

CONST ADDomain = "DC=ad,DC=joii,DC=org"
CONST ADserverIP = "10.4.2.10"
CONST ADserverIP2 = "10.4.2.11"
CONST ADS_SCOPE_SUBTREE = 2
CONST sourceFilesUrl = "https://signatures.joii.org/"
CONST sourceFilesBackup = "http://10.4.2.6/"
CONST LDAPurl="LDAP://OU=Signatures,OU=Groups,OU=JOII,DC=ad,DC=joii,DC=org"
CONST LinkImages = False