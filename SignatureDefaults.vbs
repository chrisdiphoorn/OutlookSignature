' ============================================================================================
' Author:       Chris Diphoorn 
' Create date:  23/07/2020
' Description:  Sets the Domain Signature Varibles
' Engine:		Joiitech Signature Engine	
' ============================================================================================
'
Option Explicit

CONST ADDomain = "DC=ad,DC=homecorp,DC=com"
CONST ADS_SCOPE_SUBTREE = 2
DIM sourceFilesUrl : sourceFilesUrl = "https://signatures.homecorp.com/"
CONST sourceFilesIP = "http://192.168.100.74/"
CONST LDAPurl="LDAP://OU=Signatures,OU=Groups,OU=AD - Homecorp Constructions,DC=ad,DC=homecorp,DC=com"
                      