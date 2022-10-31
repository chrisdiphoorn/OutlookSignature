' ============================================================================================
' Author:       Chris Diphoorn 
' Create date:  23/07/2020
' Description:  Sets the Domain Signature Varibles
' Engine:		Joiitech Signature Engine	
' ============================================================================================
'
Option Explicit

CONST ADDomain = "DC=XX,DC=XX,DC=XX"  ' Main LDIF Distingushed Name value of the Active Directory Domain Name.

CONST ADS_SCOPE_SUBTREE = 2  ' Leave this as 2

DIM sourceFilesUrl : sourceFilesUrl = "https://signatures.XX.XX/" ' Keep the "/" on the end.

CONST sourceFilesIP = "http://XXX.XXX.XXX.XXX/"  ' Keep the "/" on the end.

CONST LDAPurl = "LDAP://OU=Signatures,OU=Groups,OU=XXXXXXX,DC=XX,DC=XX,DC=XX"  ' OU location for any Signature Groups
                      
