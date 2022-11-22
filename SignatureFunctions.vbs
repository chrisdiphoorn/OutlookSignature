' ============================================================================================
' Author:       Chris Diphoorn 
' Create date:  04/12/2019
' Description:  Create an Users Outlook Email Signature using details from Activedirectory
'				and a .TPL HTML file.
' Engine:		Joiitech Signature Engine	
' ============================================================================================
'
Const ADS_NAME_INITTYPE_DOMAIN = 1
Const ADS_NAME_INITTYPE_SERVER = 2
Const ADS_NAME_INITTYPE_GC = 3

Const ADS_NAME_TYPE_1779 = 1
Const ADS_NAME_TYPE_CANONICAL = 2
Const ADS_NAME_TYPE_NT4 = 3
Const ADS_NAME_TYPE_DISPLAY = 4
Const ADS_NAME_TYPE_DOMAIN_SIMPLE = 5
Const ADS_NAME_TYPE_ENTERPRISE_SIMPLE = 6
Const ADS_NAME_TYPE_GUID = 7
Const ADS_NAME_TYPE_UNKNOWN = 8
Const ADS_NAME_TYPE_USER_PRINCIPAL_NAME = 9
Const ADS_NAME_TYPE_CANONICAL_EX = 10
Const ADS_NAME_TYPE_SERVICE_PRINCIPAL_NAME = 11
Const ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME = 12

Const HKEY_CLASSES_ROOT   = &H80000000
Const HKEY_CURRENT_USER   = &H80000001
Const HKEY_LOCAL_MACHINE  = &H80000002
Const HKEY_USERS          = &H80000003

Const REG_SZ        = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY    = 3
Const REG_DWORD     = 4
Const REG_MULTI_SZ  = 7

DIM NON_DOMAIN_USER: NON_DOMAIN_USER = ""
DIM NON_DOMAIN_PASSWORD: NON_DOMAIN_PASSWORD = ""

DIM DisplayNone : DisplayNone = "display: none;"
DIM DisplayHidden: DisplayHidden = " class=hidden"
DIM objOU, UseDebug, DebugFile 
DIM SignatureGroup(50), SignatureGroupFileName(50), SignatureGroupName(50), MaxSignatureGroup, aSignatureGroupName
DIM MemberGroupName(100), MemberGroupDescription(100)
DIM readreplaceemail
Dim use_signature_image_folders: use_signature_image_folders = False
Dim use_debug: use_debug = False
DIM tmpDebug
Dim name: name = " "
Dim name_v 
Dim firstname :firstname =" "
Dim firstname_v
Dim lastname : lastname =" "
Dim lastname_v
Dim title: title = " "
Dim title_v 
Dim phone: phone = " "
Dim phone_v 
Dim mobile: mobile = " "
Dim mobile_v 
Dim ipphone: ipphone =" "
Dim ipphone_v
Dim email: email = " "
Dim email_v 
Dim address: address = " "
Dim address_v 
Dim city: city = " "
Dim city_v 
Dim state: state = " "
Dim state_v 
Dim country: country = " "
Dim country_v 
Dim postcode: postcode = " "
Dim postcode_v
Dim pobox: pobox =" "
Dim pobox_v 
Dim suburb: suburb = " "
Dim suburb_v 
dim office: office =" "
dim office_v 
dim webpage: webpage =" "
dim href
dim webpage_v 
dim countryname: countryname =" "
dim countryname_v
dim department: department= " "
dim department_v 
dim notes: notes = " "
dim notes_v
Dim company: company = " "
Dim company_v, companyphone, companyphone_v
Dim companyfax,companyfax_v 
Dim companyurl, companyurl_v
dim extranote, extranote_v, extranote_t
dim footertext, footertext_v
DIM socialicons_v

Dim SignatureWidth, SignatureHeight
DIM IDnumber, IDnumber_v, userName
DIM vcardPhoto, FullStateName
DIM membernote
DIM Spc: Spc="&nbsp;"
DIM CountryNames, StateNames
DIM CountryCodes, StateCodes
DIM InternationalPrefix 
CountryNames = Array("Afghanistan","Albania","Algeria","Andorra","Angola","Anguilla","Antigua and Barbuda","Argentina","ARIPO","Armenia","Aruba","Australia","Austria","Azerbaijan","Bahamas","Bahrain","Bangladesh","Barbados","BBM and BBDM","Belarus","Belgium","Belize","Benin","Bermuda","Bhutan","Bolivia","Bosnia and Herzegovina","Botswana","Bouvet Island","Brazil","British Virgin Islands","Burundi","Brunei Darussalam","Bulgaria","Burkina Faso","Burma","Cameroon","Canada","Cape Verde","Cayman Islands","Central African Republic","Chad","Chile","China","Colombia","Comoros","Congo","Cook Islands","Costa Rica","Cote d’Ivoire","Croatia","Cuba","Cyprus","Czech Republic","Czechoslovakia","Denmark","Djibouti","Dominica","Dominican Republic","EAPO","East Germany","East Timor","Ecuador","Egypt","El Salvador","EPO","Equatorial Guinea","Estonia","Ethiopia","Falkland Island","Faroe Islands","Fiji","Finland","France","Gabon","Gambia","Greenland","Georgia","Germany","Ghana","Gibraltar","Greece","Grenada","Guatemala","Guinea","Guinea-Bissau","Guyana","Haiti","Honduras","Hong Kong","Hungary","Iceland","IIB","India","Indonesia","Iran","Iraq","Ireland","Israel","Italy","Jamaica","Japan","Jordan","Kampuchea","Kazakhstan","Kenya","Kiribati","Kuwait","Kyrgyzstan","Laos","Lebanon","Lesotho","Liberia","Libya","Liechtenstein","Lithuania","Luxembourg","Macao","Macedonia","Madagascar","Malawi","Malaysia","Maldives","Mali","Malta","Marshall Islands","Mauritania","Mauritius","Mexico","Micronesia","Moldova","Monaco","Mongolia","Montserrat","Morocco","Mozambique","Myanmar","Namibia","Nauru","Nepal","Netherlands(Holland)","Netherlands Antilles","New Zealand","Nicaragua","Niger","Nigeria","North Korea","Norway","Northern Mariana Islands","OAPI","OHIM","Oman","Other Country","Other Country in Africa","Other Country in Asia","Other Country in Europe","Other Country in North America","Other Country In Oceania","Other Country in South America","Pakistan","Panama","Papua New Guinea","Paraguay","Peru","Philippines","Poland","Portugal","Qatar","Republic of Yemen","Rhodesia","Romania","Russian Federation","Rwanda","Saint Helena","Saint Kitts and Nevis","Saint Lucia","Saint Vincent and the Grenadines","Samoa","San Marino","Sao Tome and Principe","Saudi Arabia","Senegal","Seychelles","Sierra Leone","Singapore","Slovakia","Slovenia","Solomon Islands","Somalia","South Africa","South Georgia and The South Sandwich Islands","South Korea","Spain","Sri Lanka","Sudan","Suriname","Swaziland","Sweden","Switzerland","Syria","Taiwan","Tajikistan","Tanzania","Thailand","Togo","Tonga","Trinidad and Tobago","Tunisia","Turkey","Turkmenistan","Turks and Caicos Islands","Tuvalu","Uganda","Ukraine","United Arab Emirates","United Kingdom","United States of America","Uruguay","Soviet Union","Uzbekistan","Vanuatu","Vatican City","Venezuela","Vietnam","Volta","Western Sahara","WIPO","Yemen","Yugoslavia","Zaire","Zambia","Zimbabwe")
CountryCodes = Array("AF","AL","DZ","AD","AO","AI","AG","AR","AP","AM","AW","AU","AT","AZ","BS","BH","BD","BB","BX","BY","BE","BZ","BJ","BM","BT","BO","BA","BW","BV","BR","VG","BI","BN","BG","BF","BU","CM","CA","CV","KY","CF","TD","CL","CN","CO","KM","CG","CK","CR","CI","HR","CU","CY","CZ","CS","DK","DJ","DM","DO","EA","DD","TP","EC","EG","SV","EP","GQ","EE","ET","FK","FO","FJ","FI","FR","GA","GM","GL","GE","DE","GH","GI","GR","GD","GT","GN","GW","GY","HT","HN","HK","HU","IS","IB","IN","ID","IR","IQ","IE","IL","IT","JM","JP","JO","KH","KZ","KE","KI","KW","KG","LA","LB","LS","LR","LY","LI","LT","LU","MO","MK","MG","MW","MY","MV","ML","MT","MH","MR","MU","MX","FM","MD","MC","MN","MS","MA","MZ","MM","NA","NR","NP","NL","AN","NZ","NI","NE","NG","KP","NO","MP","OA","EM","OM","XX","X6","X1","X4","X2","X5","X3","PK","PA","PG","PY","PE","PH","PL","PT","QA","YD","RH","RO","RU","RW","SH","KN","LC","VC","WS","SM","ST","SA","SN","SC","SL","SG","SK","SI","SB","SO","ZA","GS","KR","ES","LK","SD","SR","SZ","SE","CH","SY","TW","TJ","TZ","TH","TG","TO","TT","TN","TR","TM","TC","TV","UG","UA","AE","GB","US","UY","SU","UZ","VU","VA","VE","VN","HV","EH","WO","YE","YU","ZR","ZM","ZW")
StateNames = Array("Queensland","Victoria","New South Wales","South Australia","Tasmania","Australian Capital Territory","Jervis Bay Territory", "Northan Territory")
StateCodes = Array("QLD","VIC","NSW","SA","TAS","ACT","JBT","NT")
		
DIM fileSystem: Set fileSystem = CreateObject( "Scripting.FileSystemObject" )
DIM shell: Set shell = CreateObject( "WScript.Shell" )
DIM signaturesFolderPath: signaturesFolderPath = shell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Signatures\"
DIM signaturesFolderPathRoot: signaturesFolderPathRoot = shell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\"

Dim template, templateFileName, xmastext_v, xmastext, imagename, xmasimagename, xmastext_t
Dim SignatureName, EmailDomain
Dim HighlightColor, XMASColor, aditionalimage
Dim currentUser, currentUserGroups
Dim WhenChanged
Dim scriptLocation: scriptLocation = fileSystem.GetParentFolderName(WScript.ScriptFullName)
Dim adSystemInfo: Set adSystemInfo = CreateObject("ADSystemInfo")

Dim DefaultCompany, DefaultFax, DefaultPhone, DefaultDomain , DefaultEmailDomain, DefaultPOBox, DefaultName, DefaultEmail, DefaultPostCode, DefaultSuburb, DefaultWWW,DefaultHREF
DIM DefaultTitle, DefaultAddress, DefaultState, DefaultMobile, DefaultCountry
DIM ADAddress, ADTitle, ADCompany, ADPhone, ADMobile, ADstate, ADpostcode
DIM ccompany, ccompanyfax, ccompanyphone
DIM uuseraddress, uuserstate, uusersuburb, uuserpobox, uusermobile, uuserpostcode
DIM cyear,cmonth, cday, cdayname, cmonthname, cmmonth, cdday

DIM DataChanged: DataChanged = False
DIM DataChangedDate: DataChangedDate =""

DIM SocialIcon(10), SocialIconLink(10), SocialIconAlt(10), SocialIconName, Socialtmpname, ForceSocialIcon
DIM MaxSocialIcons: MaxSocialIcons = 0
DIM FoundSocialIcon, FoundSocialIconLink, FoundSocialIconAlt
DIM ForceCreate, ForceSignature
DIM AddPhoneSpace, AddAddressSpace, AddAdditionalCitySpace, AddAdditionalSuburbSpace, AddAdditionalStateSpace, AddAdditionalPOBoxSpace
DIM AddAdditionalAddressSpace, AddAdditionalTitleSpace, AddAdditionalPrevTitleSpace,CheckAdditionalSpace
DIM DisplaySocialIcons
DIM DefaultImageType
DIM TransformTitle, TransformName, TransformState, TransformPhone, TransformMobile, TransformEmail, TransformCity, TransformCountry, TransformCompanyPhone
DIM AutoPOBox
DIM Variable(100), VariableValue(100)
DIM MaxVariables
DIM InsertHTML(100), InsertHTMLCode(100), InsertHTMLFind(100), MaxInsertHTML, FindInsertHTML, FindInsertHTMLCode, FindInsertHTMLFind, tmpInsertHTML
DIM EncodeImage(20), EncodeImageData(20), EncodeImageVariable(20), MaxEncodeImage, FindEncodeImage
DIM tmpPhone, tmpCompanyPhone
DIM RemoveCompanyPhone, AdminDefaultSignature
DIM OverrideTitle, OverrideSignatureTitle, OverrideAddress, OverrideState, OverridePostCode, OverrideCompanyPhone, OverrideEmail
DIM Combine(20), MaxCombine, FindCombine, FindCombineHTML, SplitCombine, FindCombineField, FindCombineLast, CombineIfBlank,CombineErr
DIM CombineField(20,20), CombineHTML(20,20), MaxCombineHTML(20), CombineLastCode(20), MaxFindCombineHTML, CombineBlankCode(20)
DIM HideTableRows, IncludeFile, IncludeFileHTML
DIM LinkWWWColor, SymbolColor, TextColor, TextColorHighlight, TextFooterColor, BarColor, HyperLinkColor, DefaultFont, DefaultFontSize, LargerFontSize, DefaultLineHeight
DIM TestXmas, XmasOverWriteColor, xmashighlight, xmasoverwrite, xmassignature, IsXmas, AutoXmas, FirstMonday, TestXmasDate, XmasStartYear, XmasFinishYear, XmasDay, XmasAdjust
DIM findTitle, findAddress, findSignature, checkpos, FindName
DIM ExtraFile(50), MaxExtraFile, FindExtraFile
DIM XmasImageSwap(100), MaxXmasImageSwap, FindImageSwap 
DIM ManualUser

MaxSignatureGroup = 0
ManualUser = False


Function GetSignatureGroups (aSignatureGroupName)
	DIM A, objGroup
    on error resume next
	Set objOU = GetObject(aSignatureGroupName)
	A = 0
	if err.number = 0 then 
		on error goto 0
		objOU.Filter = Array("Group")
		AddDebug "Checking for AD Groups in : "& aSignatureGroupName
		For Each objGroup in objOU
			A = A + 1
			SignatureGroup(A) = objGroup.distinguishedName
			SignatureGroupFileName(A) = objGroup.description
			SignatureGroupName(A) = replace(objGroup.Name,"CN=","")
			AddDebug SignatureGroup(A) & " - " & SignatureGroupFileName(A) & " - " & SignatureGroupName(A)
		Next      
		AddDebug "Groups Found : " & cstr(A)
	end if
	GetSignatureGroups = A
End Function

SUB DeleteSignature(SignatureName)
		'SignatureImageFolder
    if len(SignatureName) > 0 then 
        DeleteSignatureFile(signaturesFolderPath & signatureName & ".htm")
        DeleteSignatureFile(signaturesFolderPath & signatureName & ".tpl")
		DeleteSignatureFile(signaturesFolderPath & signatureName & ".txt")
		DeleteSignatureFile(signaturesFolderPath & signatureName & ".jpg")
		DeleteSignatureFile(signaturesFolderPath & signatureName & ".bmp")
		DeleteSignatureFile(signaturesFolderPath & signatureName & ".png")
		DeleteSignatureFile(signaturesFolderPath & signatureName & ".vcf")
		DeleteSignatureFile(signaturesFolderPath & signatureName & "-Youtube.png")
		DeleteSignatureFile(signaturesFolderPath & signatureName & "-Instagram.png")
		DeleteSignatureFile(signaturesFolderPath & signatureName & "-LinkedIn.png")
		DeleteSignatureFile(signaturesFolderPath & signatureName & "-Facebook.png")
		DeleteSignatureFile(signaturesFolderPath & signatureName & "-Xmas.png")
		DeleteSignatureFile(signaturesFolderPath & signatureName & "-Xmas.jpg")
		
		if use_signature_image_folders = True then 
		
		end if
    end if

END SUB

Sub DeleteSignatureFile(byval htmlSignatureFile)

    on error resume next
    IF fileSystem.FileExists(htmlSignatureFile) THEN 
		filesystem.DeleteFile(htmlSignatureFile)
		AddDebug "Deleting Signature File : " & htmlSignatureFile
	END IF		
    on error goto 0

END SUB


Sub CreateSignature(SignatureName, templateFileName)


	' Dont run this proceedure if the name value is blank.
	IF SignatureName ="" or templateFileName = ".tpl" then exit sub
    
	'templateFileName = SignatureName +".tpl"
	
	' Cleanup values so they cant be re-used in other signatures by accident.
	HighlightColor="": aditionalimage="": xmastext_v="": xmastext="": imagename="":EmailDomain=""
	DefaultCompany="":DefaultFax="":DefaultPhone="":DefaultDomain="":DefaultEmailDomain="":DefaultPOBox="":DefaultCountry="":DefaultEmail="":DefaultSuburb="":DefaultWWW="":DefaultHREF=""
	DefaultTitle="": DefaultName = "": DefaultState ="":DefaultAddress="": DefaultMobile="": DefaultPostCode =""
	ccompany="":ccompanyfax="":ccompanyphone="":Company="":CompanyFax="":CompanyPhone=""
	uuseraddress="": uuserstate ="":uusersuburb="":uuserpobox ="":uusermobile="":uuserpostcode=""
	SignatureWidth = "": SignatureHeight = ""
	companyurl="": readreplaceemail ="":SocialIconName="":ForceSocialIcon=""
    vcardPhoto ="": FullStateName="": extranote ="":extranote_v ="": membernote =""
	TransformTitle="":TransformState="":TransformMobile="": TransformName="": TransformPhone="": TransformCity="":TransformCountry=""
	ForceSignature = False: ForceCreate=False: MaxSocialIcons =0
	AddPhoneSpace = False: AddAddressSpace = False
	AddAdditionalCitySpace = False: AddAdditionalSuburbSpace= False: AddAdditionalPOBoxSpace = False: AddAdditionalStateSpace = False
	AddAdditionalAddressSpace = False: AddAdditionalTitleSpace = False: AddAdditionalPrevTitleSpace = False
	AutoPOBox = False
	MaxVariables = 0: MaxInsertHTML = 0: footertext =""
	MaxEncodeImage=0: InternationalPrefix ="": DisplaySocialIcons = True
	RemoveCompanyPhone = False:	AdminDefaultSignature = False
	OverrideTitle ="": OverrideSignatureTitle="": OverrideAddress = "": OverrideState ="": OverridePostCode="":OverrideCompanyPhone="":OverrideEmail=""
	HideTableRows = False: IncludeFile="": IncludeFileHTML=""
	LinkWWWColor="":SymbolColor="":TextColor="":TextColorHighlight="":TextFooterColor="":BarColor="":HyperLinkColor="":DefaultFont="Lucida,sans-serif":DefaultFontSize="11.0pt":LargerFontSize="11.0pt": DefaultLineHeight="11.0pt"
	TestXmas = False: XmasOverWriteColor ="":xmashighlight ="": xmasoverwrite ="": xmassignature =""
	findTitle="": findAddress ="": findAddress ="": findSignature =""
	IsXmas = false: AutoXmas = False: TestXmasDate ="": XmasDay = "": XmasAdjust = 1

    DIM xmasfrom, xmasto, part1, p1, NewMonday, XmasLastDay, XmasNextLastDay, preMonday, NewYearStartWeek, WeeksClosed
    DIM SignatureImageFolder, templateFilePath, headerHTML, templateHTML,  templateTEXT
	DIM PreviousInfoFilePath , tmpstring
	
	' Temporary Turn off Debug to get the Create Details
	use_debug = True
	
	IF ManualUser = True then tmpstring="Manually "
	
	AddDebug ""
	AddDebug "************************************************************************************************************************************"
	AddDebug "* " & Date & " * " & tmpstring &"Creating Signature " & SignatureName & " using file " & templateFileName
	AddDebug "************************************************************************************************************************************"
    
    ' Create signature folder if it doesn't exist    
	IF fileSystem.FolderExists(signaturesFolderPathRoot) THEN
		IF NOT fileSystem.FolderExists(signaturesFolderPath) THEN
			AddDebug "Creating Users Outlook Signature Folder : " & signaturesFolderPath
			on error resume next
			fileSystem.CreateFolder(signaturesFolderPath)
			on error goto 0		   
		END If
	END IF
	
	' Create Signature Image Folders if the option is set to true
    IF use_signature_image_folders = True THEN
	    SignatureImageFolder="\"&SignatureName&"_files"
        IF NOT fileSystem.FolderExists(signaturesFolderPath&SignatureImageFolder) THEN
	       AddDebug "Creating Signature Image Folder : " & signaturesFolderPath&SignatureImageFolder
		   on error resume next
           fileSystem.CreateFolder(signaturesFolderPath&SignatureImageFolder)
           on error goto 0
        END IF
    END IF

	' Get the Standard Users values from Activedirectory if not using the _Signatures.INI File
	IF ManualUser <> True then 
		UpdateGlobalVarsFromAD() 
	END IF
	
    ' Get the Current Date Details
	cyear = cint(year(date))
    cmonth = cint(month(date))
	cday = cint(day(date))
	
	' Get the current Date Vales with 2 Digits
	cdday = trim(cstr(cday))
	if len(cdday) =1 then cdday ="0"+cdday
	cmmonth = trim(cstr(cmonth))
	if len(cmmonth) =1 then cmmonth ="0"+cmmonth
	
    cdayname = WeekdayName(Weekday(date, 1),False,1)
    cmonthname = MonthName(cmonth)

    IF ManualUser = True then 
		IF filesystem.fileexists(signaturesFolderPath + templateFileName) THEN 
			AddDebug "Using Local FILE " & templateFileName & " FROM " & signaturesFolderPath
			templateFilePath = signaturesFolderPath + templateFileName
		END IF
	ELSE
		templateFilePath = DownloadFile(templateFileName, sourceFilesUrl, signaturesFolderPath)
	end if
	
	' Turn off Debug again and it wil be enabled again if found in the .tlp file.
	use_debug = False
	
	' If there is a Template file (.tpl) then use it 	
	IF LEN(templateFilePath) > 0 THEN
    
		templateTEXT=""
		templateHTML=""
		headerHTML=""
	
		templateHTML = GetOutlookSignatureHtml(templateFilePath)
		headerHTML = ReadHeaderSignatureHtml(templateHTML, SignatureName)

		IF LEN(headerHTML) >0 THEN
		
			if instr(lcase(headerHTML),"<*|debug|*true>") > 0 then 
				use_debug = True
			else
				use_debug = False
			end IF
	
			AddDebug "Found headerHTML settings."
	   
			TestXmasDate = ReadDefaultSettings("TestXmas", headerHTML)
	   		
			IF len(TestXmasDate) > 0 Then
				AddDebug "**** Testing Xmas Enabled! **** " 
				if ucase(TestXmasDate) = "TRUE" or ucase(TestXmasDate) = "FALSE" THEN 
				TestXmasDate = ""
			ELSE
				TestXmasDate = ConvertToDate(TestXmasDate)
				if len(TestXmasDate) = 10 then 
					cyear = int(right(TestXmasDate,4))
					cmonth = int(mid(TestXmasDate,4,2))
					cday = int(left(TestXmasDate,2))
					AddDebug "Changing Current Date to " & TestXmasDate
				Else
					AddDebug "*** Invalid Xmas Test Date: " & TestXmasDate 
				end if
			END IF
			TestXmas = True
		END IF
		
		IncludeFileHTML = ""
	    IncludeFile = ReadDefaultSettings("Include", headerHTML)
		if len(IncludeFile) > 0 then
			AddDebug "Downloading Include File :" & sourceFilesUrl+IncludeFile
			IncludeFileHTML= GetOutlookSignatureHtml(sourceFilesUrl+IncludeFile)
			if len(IncludeFileHTML) > 0 then 
				tmpstring=instr(templateHTML,"-->")
				if tmpstring > 0 then 
					templateHTML = LEFT(templateHTML,tmpstring + 3) + IncludeFileHTML + MID(templateHTML, tmpstring + 3, LEN(templateHTML))
				end if
			end if
	    end if
		
	    DefaultImageType = lcase(ReadDefaultSettings("DefaultImageType", headerHTML))
		if len(DefaultImageType) = 0 then DefaultImageType = "jpg"
		
		companyurl = ReadDefaultSettings("CompanyURL", headerHTML)
		HighlightColor = ReadDefaultSettings("HighlightColor", headerHTML)
		imagename = ReadDefaultSettings("ImageName", headerHTML)
		aditionalimage = ReadDefaultSettings("AditionalImage", headerHTML)
		SignatureWidth = ReadDefaultSettings("Width", headerHTML)
		SignatureHeight = ReadDefaultSettings("Height", headerHTML)
		
		InternationalPrefix = ReadDefaultSettings("InternationalPrefix",headerHTML)
		
		ccompany = ReadDefaultSettings("Company", headerHTML)
		ccompanyfax = ReadDefaultSettings("CompanyFax", headerHTML)
		ccompanyphone = ReadDefaultSettings("CompanyPhone", headerHTML)
		emaildomain = ReadDefaultSettings("EmailDomain", headerHTML)
		
		uuserpobox = ReadDefaultSettings("UserPOBox", headerHTML)
		uuseraddress = ReadDefaultSettings("UserAddress", headerHTML)
		uuserstate = ReadDefaultSettings("UserState", headerHTML)
		uusersuburb = ReadDefaultSettings("UserSuburb", headerHTML)
		uusermobile = ReadDefaultSettings("UserMobile", headerHTML)
		uuserpostcode = ReadDefaultSettings("UserPostCode", headerHTML)
		
		vcardPhoto = ReadDefaultSettings("vCardPhoto", headerHTML)
		membernote = ReadDefaultSettings("MemberNote", headerHTML)
		
		readreplaceemail = ReadDefaultSettings("ReplaceEmailDomain",headerHTML)
		if len(readreplaceemail) > 0 and len(email) > 0 then 
			email = ReplaceEmailDomain(email, readreplaceemail)
		end if
		
		' These settings if defined are using another value if one does not exist (Rendra vocabulary - Press 1**) - No Debug output
		if len(ccompany) and len(trim(company)) = 0 then company = ccompany
		if len(ccompanyfax) and len(trim(companyfax)) = 0 then companyfax = ccompanyfax
		if len(ccompanyphone) and len(trim(companyphone)) = 0 then companyphone = ccompanyphone

		if len(uuserpobox) and len(trim(pobox)) = 0 then pobox = uuserpobox
		if len(uuseraddress) and len(trim(address)) = 0 then address = uuseraddress
		if len(uusermobile) and len(trim(mobile)) = 0 then mobile = uusermobile
		if len(uuserstate) and len(trim(state)) = 0 then state = uuserstate
		if len(uusersuburb) and len(trim(suburb)) = 0 then suburb = uusersuburb
		if len(uuserpostcode) and len(trim(postcode)) = 0 then postcode = uuserpostcode
		
		' These settings if defined are overriding the value - can even blank a value out.
		    DefaultCompany = ReadDefaultSettings("DefaultCompany", headerHTML)
		        DefaultFax = ReadDefaultSettings("DefaultFax", headerHTML)
		      DefaultPhone = ReadDefaultSettings("DefaultPhone", headerHTML)
		DefaultEmailDomain = ReadDefaultSettings("DefaultEmailDomain",headerHTML)
		      DefaultEmail = ReadDefaultSettings("DefaultEmail",headerHTML)
		       DefaultName = ReadDefaultSettings("DefaultName",headerHTML)
		      DefaultTitle = ReadDefaultSettings("DefaultTitle",headerHTML)
		      DefaultState = ReadDefaultSettings("DefaultState",headerHTML)
		    DefaultAddress = ReadDefaultSettings("DefaultAddress",headerHTML)
		     DefaultSuburb = ReadDefaultSettings("DefaultSuburb",headerHTML)
		     DefaultMobile = ReadDefaultSettings("DefaultMobile",headerHTML)
		    DefaultCountry = ReadDefaultSettings("DefaultCountry",headerHTML)
		      DefaultPOBox = ReadDefaultSettings("DefaultPOBox",headerHTML)
		   DefaultPostCode = ReadDefaultSettings("DefaultPostCode",headerHTML)	
		        DefaultWWW = ReadDefaultSettings("DefaultWWW",headerHTML)	
		       DefaultHREF = ReadDefaultSettings("DefaultHREF",headerHTML)	

	        TransformTitle = ReadDefaultSettings("TransformTitle",headerHTML)
	         TransformName = ReadDefaultSettings("TransformName",headerHTML)
	        TransformState = ReadDefaultSettings("TransformState",headerHTML)
	 TransformCompanyPhone = ReadDefaultSettings("TransformCompanyPhone",headerHTML)
	        TransformPhone = ReadDefaultSettings("TransformPhone",headerHTML)
		   TransformMobile = ReadDefaultSettings("TransformMobile",headerHTML)
		    TransformEmail = ReadDefaultSettings("TransformEmail",headerHTML)
		     TransformCity = ReadDefaultSettings("TransformCity",headerHTML)
		  TransformCountry = ReadDefaultSettings("TransformCountry",headerHTML)
		  
		       xmasLastDay = ReadDefaultSettings("XmasLastDay", headerHTML)
				  xmasfrom = ReadDefaultSettings("XmasFrom", headerHTML)
		  		    xmasto = ReadDefaultSettings("XmasTo", headerHTML)
          NewYearStartWeek = ReadDefaultSettings("NewYearStartWeek", headerHTML)
			   WeeksClosed = ReadDefaultSettings("WeeksClosed", headerHTML)
			
		      LinkWWWColor = ReadDefaultSettings("LinkWWWColor",headerHTML)
		       SymbolColor = ReadDefaultSettings("SymbolColor",headerHTML)
		         TextColor = ReadDefaultSettings("TextColor",headerHTML)
		TextColorHighlight = ReadDefaultSettings("TextColorHighlight",headerHTML)
		   TextFooterColor = ReadDefaultSettings("TextFooterColor",headerHTML)
		          BarColor = ReadDefaultSettings("BarColor",headerHTML)
		    HyperlinkColor = ReadDefaultSettings("HyperlinkColor",headerHTML)
		       DefaultFont = ReadDefaultSettings("DefaultFont",headerHTML)
		 DefaultLineHeight = ReadDefaultSettings("DefaultLineHeight",headerHTML)
		   DefaultFontSize = ReadDefaultSettings("DefaultFontSize",headerHTML)
		    LargerFontSize = ReadDefaultSettings("LargerFontSize",headerHTML)
			
		if len(xmasfrom) > 0 and xmasfrom <> "DD/MM/YYYY" and xmasfrom <> "DD/12/YYYY" then 
			xmasfrom = ConvertToDate(xmasfrom)
			if len(XmasLastDay) = 0 then XmasLastDay = xmasfrom
		Else
			if cmonth =1 then 
				xmasfrom = FindStartDayinMonth("01/12/"+cstr(cyear-1))
			else
				xmasfrom = FindStartDayinMonth("01/12/"+cstr(cyear))
			end if
		end if
		
		' Convert XmasTo to a Date Format
		if len(xmasto) > 0 then 
			xmasto = ConvertToDate(xmasto)
		END IF
		
		' Convert XmasLastday to a Date Format or Set it as an Auto Date in December
		IF len(XmasLastDay) > 0 THEN
		   XmasLastDay = ConvertToDate(XmasLastDay)
		ELSE
			XmasLastDay="DD/12/YYYY"
		END IF
		
		IF cmonth =1 THEN
			XmasDay = "25/12/"+cstr(cyear-1)
		Else
			XmasDay = "25/12/"+cstr(cyear)
		END IF
		
		
		
		if right(ucase(XmasLastDay),4) = "YYYY" then
				' Get previous year if current month is January
			if cmonth =1 then 
				XmasLastDay = left(XmasLastDay,len(XmasLastDay)-4) + cstr(cyear-1)
			Else
				XmasLastDay = left(XmasLastDay,len(XmasLastDay)-4) + cstr(cyear)	
			END IF
					
		end if
		
		' Auto ajust XmasLastDay Day depending on what day it falls on.
		' What Day is Christmas Day EG: If is a sunday, then Make the last work day for the year Wednesday.
		
		IF instr(ucase(XmasLastDay),"DD") > 0 then
		
			' Just take a good guess and automatically workout the last day.	
			if Weekday(cdate(XmasDay)) =  1 then XmasAdjust = 3 ' Sunday so LastDay Wednesday
			if Weekday(cdate(XmasDay)) =  2 then XmasAdjust = 4 ' Monday so LastDay Thursday
			if Weekday(cdate(XmasDay)) =  3 then XmasAdjust = 4 ' Tuesday so LastDay Friday
			if Weekday(cdate(XmasDay)) =  4 then XmasAdjust = 2 ' Wednesday so LastDay Monday
			if Weekday(cdate(XmasDay)) =  5 then XmasAdjust = 2 ' Thursday so LastDay Tuesday
			if Weekday(cdate(XmasDay)) =  6 then XmasAdjust = 2 ' Friday so LastDay Wednesday
			if Weekday(cdate(XmasDay)) =  7 then XmasAdjust = 3 ' Saturday so LastDay Wednesday
			
			AddDebug "Set Variable: XmasAdjust -> " & cstr(XmasAdjust) & " (Minus these day(s) from 25/12)"
			
			XmasLastDay = Replace(XmasLastDay, "DD", cstr(25-XmasAdjust))
			' an alternative XmasLastDay = FindLastDay(XmasLastDay,1)
			
		END IF
		
		AddDebug "Set Variable: XmasLastDay -> " & XmasLastDay & " (Last Work Day in December)"
		
		' Check xmas details if they are being used in the signature
        if len(xmasfrom) > 0 and len(xmasto) > 0 then 
			
			'AddDebug "Set Variable: xmasfrom -> "& xmasfrom
			'AddDebug "Set Variable: xmasto 'XMAS' -> " & xmasto 
			' Christmas Period is defined sometime from Dec to Jan (Next Year)
			' Auto set the Year if the from date year is not a specific year
			if right(ucase(xmasfrom),4) = "YYYY" then
				' Get previous year if current month is January
				if cmonth =1 then 
					xmasfrom = left(xmasfrom,len(xmasfrom)-4) + cstr(cyear-1)
				ELSE
					xmasfrom = left(xmasfrom,len(xmasfrom)-4) + cstr(cyear)
				END IF
				AddDebug "Set Variable: XmasFrom -> " & xmasfrom & " (Auto Xmas Year)"
				AutoXmas = True
			end if

			IF instr(ucase(xmasfrom),"MM") > 0 then
				xmasfrom = Replace(xmasfrom, "MM", "12")
				AddDebug "Set Variable: XmasFrom -> " & xmasfrom & " (Auto Xmas Month - December)"
				AutoXmas = True
			END IF
			
			IF instr(ucase(xmasfrom),"DD") > 0 then
				xmasfrom = Replace(xmasfrom, "DD", "01")
				xmasfrom = FindStartDayinMonth(xmasfrom)
				AddDebug "Set Variable: XmasFrom -> " & xmasfrom & " (Auto - First Work Day in December)"
				AutoXmas = True
			END IF
			
			' Auto set the next year if the xmas date value does not specify a specific year
			if right(ucase(xmasto),4) = "YYYY" then 
				' Automtatically set this to the next year
				if cmonth = 1 then 
					xmasto = left(xmasto,len(xmasto)-4) + cstr(cyear)
				Else
					xmasto = left(xmasto,len(xmasto)-4) + cstr(cyear+1)
				end if
				
				AddDebug "Set Variable: XmasTo -> " & xmasto&" (Auto Year)"
				AutoXmas = True
			end if
			
			IF instr(ucase(xmasto),"MM") > 0 then
				xmasto = Replace(xmasto, "MM", "01")
				AddDebug "Set Variable: XmasTo -> " & xmasto &" (Auto Month - January)"
				AutoXmas = True
			END IF
			
			IF len(WeeksClosed) > 0 then 
				' Override the xmasto so the xmasto value is x weeks from the last week in Dec
				
			END IF
			 
			IF instr(ucase(xmasto),"DD") > 0 then
				xmasto = Replace(xmasto, "DD", "01")
				SELECT CASE NewYearStartWeek
						CASE "1"
						xmasto = FindSundayinDate(xmasto, 1)
						CASE "2"
						xmasto = FindSundayinDate(xmasto, 2)
						CASE "3"
						xmasto = FindSundayinDate(xmasto, 3)
						CASE "4"
						xmasto = FindSundayinDate(xmasto, 4)
				CASE ELSE
					xmasto = FindSundayinDate(xmasto, 2)
				END SELECT
				IF LEN(NewYearStartWeek) > 0 then 
					AddDebug "Set Variable: XmasTo -> " & xmasto & " (" & th(int(NewYearStartWeek)) & " Sunday)"
				ELSE
					AddDebug "Set Variable: XmasTo -> " & xmasto & " (Second Sunday)"
				END IF
				AutoXmas = True
			END IF
			
			If len(xmasfrom) > 0 then 
					XmasStartYear = right(xmasfrom,4)
			end if
			If len(xmasto) > 0 then 
					XmasFinishYear = right(xmasto,4)
			end if
			
			IF AutoXmas = True Then
			
				SELECT CASE NewYearStartWeek
				CASE "1"
					newMonday = FindWeekinDate(xmasto, 1)
				CASE "2"
					newMonday = FindWeekinDate(xmasto, 2)
				CASE "3"
					newMonday = FindWeekinDate(xmasto, 3)
				CASE "4"
					newMonday = FindWeekinDate(xmasto, 4)
				CASE ELSE
					newMonday = FindWeekinDate(xmasto, 2)
				END SELECT

				'if len(XmasLastDay) = 0 then XmasLastDay = FindLastDay(xmasfrom,1)  ' Extra day added 
				
				preMonday = FindWeekinDate(xmasto, 1)

				if len(XmasLastDay) > 0 then    AddDebug "Set Variable: {lastday} OR *|lastday|* -> " & DateExpanded(XmasLastDay)
				if len(XmasLastDay) > 0 then    AddDebug "Set Variable: {firstholiday} OR *|firstholiday|* -> "& DateExpanded2(DateAddDay(XmasLastDay, 1))
				if len(preMonday) > 0 then      AddDebug "Set Variable: {firstmonday} OR *|firstmonday|* -> " & DateExpanded(preMonday)
				if len(NewMonday) > 0 then      AddDebug "Set Variable: {firstday} OR *|firstday|* -> " & DateExpanded(NewMonday)
				if len(XmasTo) > 0 then         AddDebug "Set Variable: {firstholiday} OR *|firstholiday|* -> "& DateExpanded2(XmasTo)
				if len(XmasStartYear) > 0 then  AddDebug "Set Variable: {XmasStartYear} OR *|XmasStartYear|* ->"& XmasStartYear
				if len(XmasFinishYear) > 0 then AddDebug "Set Variable: {XmasFinishYear} OR *|XmasFinishYear|* -> "& XmasFinishYear
			END IF
			
			' Only check for the next xmas settings if the current date is between the xmas values or Testing Enabled
			if isBetweenDate(CDate(xmasfrom), CDate(xmasto)) = True or TestXmas = True then 
				IsXmas = True
				if TestXmas = True then 
					AddDebug "************* XMAS TESTING Enabled. -> From: " & cstr(CDate(xmasfrom)) & " To: " & cstr(CDate(xmasto)) & " TestingDate <- " & cstr(cday)&"-"&cstr(cmonth)&"-"&cstr(cyear)
				ELSE
					AddDebug "The Current Date is between the Xmas Dates. From: " & cstr(CDate(xmasfrom)) & " To: " & cstr(CDate(xmasto)) & " Currently <- " & cstr(cday)&"-"&cstr(cmonth)&"-"&cstr(cyear)
				END IF
				
				xmasimagename = ReadDefaultSettings("XmasImageName", headerHTML)
				
				if len(xmasoverwrite) > 0 then 
					if lcase(xmassignature)= lcase(SignatureName) then 
						xmastext = xmasoverwrite
						AddDebug "Mod Variable: xmastext 'Override' -> '" & xmasoverwrite
					else
						xmastext= ReadDefaultSettings("XmasText", headerHTML)
					end if
				else
				   xmastext= ReadDefaultSettings("XmasText", headerHTML)
				end if 
				
				if len(xmastext) > 0 then 
					xmastext = Replace(xmastext,"*|xmasstartyear|*",cstr(XmasStartYear))
					xmastext = Replace(xmastext,"*|xmasfinishyear|*",cstr(XmasFinishYear))
					xmastext = Replace(xmastext,"{xmasstartyear}",cstr(XmasStartYear))
					xmastext = Replace(xmastext,"{xmasfinishyear}",cstr(XmasFinishYear))
				end if
				
				if instr(xmastext,"*|lastday|*") > 0 and len(XmasLastDay) > 0 then 
					xmastext = Replace(xmastext,"*|lastday|*", DateExpanded(XmasLastDay))
					AddDebug "Mod Variable: xmastext 'Update *|lastday|*' -> '" & xmastext &"'"
				END IF
				if instr(xmastext,"*|firstholiday|*") > 0 and len(XmasLastDay) > 0 then 
					xmastext = Replace(xmastext,"*|firstholiday|*", DateExpanded2(DateAddDay(XmasLastDay, 1)))
					AddDebug "Mod Variable: firstholiday 'Update *|firstholiday|*' -> '" & xmastext &"'"
				END IF
				if instr(xmastext,"*|firstmonday|*") > 0 and len(PreMonday) > 0 then 
					xmastext = Replace(xmastext,"*|firstmonday|*", DateExpanded(preMonday))
					AddDebug "Mod Variable: xmastext 'Update *|firstmonday|* -> '" & xmastext &"'"
				END IF
				if instr(xmastext,"*|firstday|*") > 0 and len(NewMonday) > 0 then 
					xmastext = Replace(xmastext,"*|firstday|*", DateExpanded(NewMonday)) 
					AddDebug "Mod Variable: xmastext 'Update *|firstday|*' -> '" & xmastext &"'"
				END IF
				if instr(xmastext,"*|resumeday|*") > 0 and len(XmasTo) > 0 then 
					xmastext = Replace(xmastext,"*|resumeday|*", DateExpanded2(XmasTo) )
					AddDebug "Mod Variable: xmastext 'Update *|resumeday|*' -> '" & xmastext &"'"
				END IF
				
				if instr(xmastext,"{firstmonday}") > 0 and len(PreMonday) > 0 then 
					xmastext = Replace(xmastext,"{firstmonday}", DateExpanded(preMonday))
					AddDebug "Mod Variable: xmastext 'Update {firstmonday} -> '" & xmastext &"'"
				END IF
				if instr(xmastext,"{firstday}") > 0 and len(NewMonday) > 0 then 
					xmastext = Replace(xmastext,"{firstday}", DateExpanded(NewMonday)) 
					AddDebug "Mod Variable: xmastext 'Update {firstday}' -> '" & xmastext &"'"
				END IF
				
				xmashighlight = ReadDefaultSettings("XmasHighlightColor", headerHTML)  
				XmasOverWriteColor = ReadDefaultSettings("XmasOverWriteColor", headerHTML)
				
				' Replace the highlight color in the HTML code with the Xmas Highlight Color
				if len(xmashighlight) then 
					templateHTML = Replace(templateHTML,"*|XmasHighlightColor|*",xmashighlight,1,-1,1)
				end if
				
				' Setting the Xmas Image Name if it was previously blank.
				if len(xmasimagename) = 0 then 
					imagename = SignatureName +"-Xmas."+DefaultImageType
					AddDebug "Set Variable: xmasimagename -> " &xmasimagename & " (Was BLANK)"
				else
					imagename = xmasimagename
					AddDebug "Set Variable: imagename -> " &xmasimagename & " (Using XmasImageName)"
				end if
			Else
				' Its Not Christmas so remove the HTML from the Signature
				templateHTML = RemoveItemFromHTML(templateHTML, "<!--XmasStart-->", "<!--XmasFinish-->")
			end if
		end if
		
		' Change the email domain address to use the signatures default if defined.
		if len(DefaultEmailDomain) and len(email) then 
			email = ChangeDomain(email, DefaultEmailDomain)
			AddDebug "Mod Variable: email 'Default EmailDomain =" & DefaultEmailDomain &"' -> " & email
		end if
 
		' Change the email address to use firstname@domainname.com.au
		if lcase(ReadDefaultSettings("ForceEmailFirstName", headerHTML)) = "yes" and len(email) then 
			email = RemoveLastName(email)
			AddDebug "Mod Variable: email 'UseFIRSTNAME' -> " & email
		end if
		
		' Reset These values back to their original Active Directory User values		
		  address = ADAddress
		    title = ADTitle
		  company = ADCompany
		    phone = ADPhone
		   mobile = ADMobile
		    state = ADstate
		 postcode = ADpostcode
		 
		' Force these variables to be the default values if they have been defined.
		     company = SetToDefault("company", "DefaultCompany")
		  companyfax = SetToDefault("companyfax", "DefaultFax")
		companyphone = SetToDefault("companyphone", "DefaultPhone")
		 emaildomain = SetToDefault("emaildomain", "DefaultEmailDomain")
		       pobox = SetToDefault("pobox", "DefaultPOBox")
		       email = SetToDefault("email", "DefaultEmail")
		        name = SetToDefault("name", "DefaultName")
		       title = SetToDefault("title", "DefaultTitle")
		       state = SetToDefault("state", "DefaultState")
		     address = SetToDefault("address", "DefaultAddress")
		      suburb = SetToDefault("suburb", "DefaultSuburb")
		      mobile = SetToDefault("mobile", "DefaultMobile")
		     country = SetToDefault("country", "DefaultCountry")
	        postcode = SetToDefault("postcode","DefaultPostCode")
		     webpage = SetToDefault("webpage","DefaultWWW")
	    
		' Find the Signature Name in the Notes Field
		' The Checkpos value moves to the next part of the string after each value is obtained
		checkpos = instr(lcase(notes), "+signature(" & lcase(SignatureName)& ")" )
		if checkpos = 0 then checkpos = instr(lcase(notes), "signature(" & lcase(SignatureName)& ")" )
		if checkpos > 0 then 
			findSignature = SignatureName
			AddDebug "Found '"+ SignatureName +" 'in Notes -> @ (" + cstr(checkpos) + ")"
			checkpos = checkpos + len(findSignature)
		end if
		
		' Overrides are used for every signature file that is defined in the notes field
		if checkpos > 0 then 
			if instr(checkpos,notes,"+title(") or instr(checkpos, notes,"+title{") then
				' Modifies the title if the signatures name matches signature looks for title(xxx) or title{xxx}
				findTitle = GetText(checkpos,notes, "+title(",")","")
				if len(findTitle) = 0 then findTitle = GetText(checkpos, notes, "+title{","}","")
				if len(findTitle) > 0 then OverrideTitle = findTitle
			end if
			if instr(checkpos,notes,"+address(") or instr(checkpos, notes,"+address{") then
				' Modifies the address matches signature address(xxx) or address{xxx}
				findAddress = GetText(checkpos, notes, "+address(",")","")
				if len(findAddress) = 0 then findAddress = GetText(checkpos, notes, "+address{","}","")
				if len(findAddress) > 0 then OverrideAddress = replace(findAddress,"^",chr(13))
			end if
			if instr(checkpos,notes,"+state(") or instr(checkpos, notes,"+state{") then
				' Modifies the state if address(xxx) or address{xxx}
				findAddress = GetText(checkpos, notes, "+state(",")","")
				if len(findAddress) = 0 then findAddress = GetText(checkpos, notes, "+state{","}","")
				if len(findAddress) > 0 then OverrideState=findAddress
			end if
			if instr(checkpos,notes,"+postcode(") or instr(checkpos, notes,"+postcode{") then
				' Modifies the postcode if postcode(xxx) or postcode{xxx}
				findAddress = GetText(checkpos, notes, "+postcode(",")","")
				if len(findAddress) = 0 then findAddress = GetText(checkpos, notes, "+postcode{","}","")
				if len(findAddress) > 0 then OverridePostcode=findAddress
			end if
			if instr(checkpos,notes,"+email(") or instr(checkpos, notes,"+email{") then
				' Modifies the email address if email(xxx) or email{xxx}
				findAddress = GetText(checkpos, notes, "+email(",")","")
				if len(findAddress) = 0 then findAddress = GetText(checkpos, notes, "+email{","}","")
				if len(findAddress) > 0 then OverrideEmail=findAddress
			end if
		end if

		' Override the values with ones that have been evaluated in the users active directory notes field
		href = webpage
		If len(DefaultHREF) then href=DefaultHREF
		if RemoveCompanyPhone = True then companyphone = ""

		if len(OverrideTitle) > 0 and len(findSignature) > 0 then 
			AddDebug "Mod Variable: title 'Notes Override' -> " & OverrideTitle & " <- Previous Value: " & title
			title = OverrideTitle 
			OverrideTitle = ""
		end if 
		if len(OverrideAddress) > 0 and len(findSignature) > 0 then
			AddDebug "Mod Variable: address 'Notes Override' -> " & OverrideAddress & " <- Previous Value: " & address
			address = OverrideAddress
			OverrideAddress = ""
		end if 
		if len(OverrideState) > 0 and len(findSignature) > 0 then
			AddDebug "Mod Variable: state 'Notes Override' -> " & OverrideState & " <- Previous Value: " & state
			state = OverrideState
			OverrideState = ""
		end if 
		if len(OverridePostCode) > 0 and len(findSignature) > 0 then
			AddDebug "Mod Variable: postcode 'Notes Override' -> " & OverridePostCode & " <- Previous Value: " & postcode
			postcode = OverridePostCode
			OverridePostCode = ""
		end if 
		if len(OverrideEmail) > 0 and len(findSignature) > 0 then
			AddDebug "Mod Variable: email 'Notes Override' -> " & OverrideEmail & " <- Previous Value: " & email
			email = OverrideEmail
			OverrideEmail = ""
		end if 
		if len(OverrideCompanyPhone) then 
			AddDebug "Mod Variable: companyphone 'Override' -> " & OverrideCompanyPhone & " <- Previous Value: " & companyphone
			companyphone = OverrideCompanyPhone
			OverrideCompanyPhone = ""
		end if 
		
		' Create a unique ID number that can be placed into a Signature and used to track emails instead of using the users name.
		IDnumber = CreateIDNumber(userName)	
		
		' if companyphone and phone are the same phone number then clear the phone value
		' otherwise you will end up with the same number showing twice in the signature.
		IF len(companyphone) and len(phone) Then
			' Remove space from phone number to ensure they can match better
			tmpPhone = Replace(phone," ","")
			tmpCompanyPhone = Replace(companyphone," ","")
			if tmpPhone = tmpCompanyPhone then 
				phone = ""
				AddDebug "Mod Variable: phone 'CompanyPhone is same as phone' -> Now Blank"
			end if
		END IF
		
		' Turn on an option on using if True 
		IF ucase(ReadDefaultSettings("AddPhoneSpace", headerHTML)) = "TRUE"  then AddPhoneSpace = True
		IF ucase(ReadDefaultSettings("AddAddressSpace", headerHTML)) = "TRUE"  then AddAddressSpace = True
		IF ucase(ReadDefaultSettings("AddAdditionalCitySpace", headerHTML)) = "TRUE"  then AddAdditionalCitySpace = True
		IF ucase(ReadDefaultSettings("AddAdditionalPOBoxSpace", headerHTML)) = "TRUE"  then AddAdditionalPOBoxSpace = True
		IF ucase(ReadDefaultSettings("AddAdditionalSuburbSpace", headerHTML)) = "TRUE"  then AddAdditionalSuburbSpace = True
		IF ucase(ReadDefaultSettings("AddAdditionalStateSpace", headerHTML)) = "TRUE"  then AddAdditionalStateSpace = True
		IF ucase(ReadDefaultSettings("AddAdditionalAddressSpace", headerHTML)) = "TRUE"  then AddAdditionalAddressSpace = True
		IF ucase(ReadDefaultSettings("AddAdditionalTitleSpace", headerHTML)) = "TRUE"  then AddAdditionalTitleSpace = True
		IF ucase(ReadDefaultSettings("AddAdditionalPrevTitleSpace", headerHTML)) = "TRUE"  then AddAdditionalPrevTitleSpace = True
		IF ucase(ReadDefaultSettings("AutoPOBox", headerHTML)) = "TRUE"  then AutoPOBox = True
		IF ucase(ReadDefaultSettings("HideTableRows", headerHTML)) = "TRUE"	then HideTableRows = True
		
		IF ucase(ReadDefaultSettings("ForceDefaultSignature", headerHTML)) = "TRUE"  then ForceSignature= True
		IF ucase(ReadDefaultSettings("ForceCreate", headerHTML)) = "TRUE"  then ForceCreate= True
		
		' Find ExtraFiles - These are copied along with the SocialIcons, ExtraImage, PrimaryImage
		DO
		  FindExtraFile = ReadDefaultSettings("ExtraFile", headerHTML)
		  IF FindExtraFile <> "" Then
			MaxExtraFile = MaxExtraFile + 1
			ExtraFile(MaxExtraFile) = FindExtraFile
		  END IF
		LOOP until FindExtraFile = ""
		  
		' Find the SocialIcons and their Links
		SocialIconName = ReadDefaultSettings("SocialIconName", headerHTML)
		ForceSocialIcon = ReadDefaultSettings("SocialIconForce", headerHTML)
		DO
		  FoundSocialIcon = ReadDefaultSettings("SocialIcon", headerHTML)
		  FoundSocialIconLink = ReadDefaultSettings("SocialIconLink", headerHTML)
		  FoundSocialIconAlt = ReadDefaultSettings("SocialIconAlt", headerHTML)
		  if FoundSocialIcon <> "" and FoundSocialIconLink <> "" then 
				MaxSocialIcons=MaxSocialIcons+1
				SocialIcon(MaxSocialIcons) = FoundSocialIcon
				SocialIconLink(MaxSocialIcons) = FoundSocialIconLink
				SocialIconAlt(MaxSocialIcons)= FoundSocialIconAlt
				if len(FoundSocialIconAlt) = 0 then SocialIconAlt(MaxSocialIcons) = FoundSocialIcon
		  end if
		LOOP until FoundSocialIcon = ""
		FoundSocialIcon = "Social Icon"
		if MaxSocialIcons > 1 then FoundSocialIcon = FoundSocialIcon + "s"  'Cause it looks better in debug Rendra!
		if MaxSocialIcons > 0 then AddDebug "Found " & cstr(MaxSocialIcons) & " " & FoundSocialIcon
		
		' Find InsertHTML
		DO
		  FindInsertHTML = ReadDefaultSettings("InsertHTML", headerHTML)
		  FindInsertHTMLCode = ReadDefaultSettings("InsertHTMLCode", headerHTML)
		  FindInsertHTMLFind = ReadDefaultSettings("InsertHTMLFind", headerHTML)
		  
		  if FindInsertHTML <> "" and FindInsertHTMLCode <> "" and FindInsertHTMLFind <> "" then 
				MaxInsertHTML = MaxInsertHTML + 1
				InsertHTML(MaxInsertHTML) = FindInsertHTML
				InsertHTMLCode(MaxInsertHTML) = FindInsertHTMLCode
				InsertHTMLFind(MaxInsertHTML) = checkvar(FindInsertHTMLFind)			
		  end if
		LOOP until FindInsertHTML = ""
		FindInsertHTML = "InsertHTML Code Change"
		if MaxInsertHTML > 1 then FindInsertHTML = FindInsertHTML + "s"  'Cause it looks better in debug!
		if MaxInsertHTML > 0 then AddDebug "Found " & cstr(MaxInsertHTML) & " " & FindInsertHTML
		
		' Find ImageSwap
		MaxXmasImageSwap = 0
		Do
			FindXmasSwap = ReadDefaultSettings("XmasImageSwap", headerHTML)
			if FindXmasSwap <>"" then 
				MaxXmasImageSwap=MaxXmasImageSwap + 1
				XmasImageSwap(MaxXmasImageSwap) = FindXmasSwap
			end if
		LOOP until FindXmasSwap = "" or MaxXmasImageSwap= 100
		FindXmasSwap = "XmasImage Swap"
		if MaxXmasImageSwap > 1 then FindXmasSwap = FindXmasSwap + "s"  'Cause it looks better in debug!
		if MaxXmasImageSwap > 0 then AddDebug "Found " & cstr(MaxXmasImageSwap) & " " & FindXmasSwap
		
		' Find Combine and CombineHTML
		MaxCombine = 0
		Do
			FindCombine = ReadDefaultSettings("Combine", headerHTML)
			if FindCombine <> "" then 
				MaxCombine = MaxCombine + 1
				MaxFindCombineHTML = 0
				SplitCombine=split(FindCombine,"+")
				Combine(MaxCombine) = SplitCombine(0)
				MaxCombineHTML(MaxCombine) = 0
				FindCombineLast = ReadDefaultSettings("CombineLast", headerHTML)
				If FindCombineLast = "" then FindCombineLast = "True"
				CombineIfBlank = ReadDefaultSettings("CombineIfBlank", headerHTML)
				if CombineIfBlank = "" then CombineIfBlank = "False"
				CombineLastCode(MaxCombine) = FindCombineLast
				CombineBlankCode(MaxCombine) = CombineIfBlank
				CombineErr=False
				do
					FindCombineField = ReadDefaultSettings("CombineField", headerHTML)
					FindCombineHTML = ReadDefaultSettings("CombineHTMLCode", headerHTML)
					if FindCombineField <> "" then 
						MaxFindCombineHTML = MaxFindCombineHTML + 1
						CombineHTML(MaxCombine, MaxFindCombineHTML) = FindCombineHTML
						CombineField(MaxCombine, MaxFindCombineHTML) = FindCombineField
					else
						CombineErr = True
					END IF
				loop until MaxFindCombineHTML > ubound(SplitCombine) or CombineErr = True
				MaxCombineHTML(MaxCombine) = MaxFindCombineHTML
			end if
		LOOP until FindCombine = ""
		FindCombine = "Combine Field"
		if MaxCombine > 1 then FindCombine = FindCombine + "s"  'Cause it looks better in debug!
		if MaxCombine > 0 then AddDebug "Found " & cstr(MaxCombine) & " " & FindCombine
		if CombineErr = True then 
			AddDebug "Could not find all the Combine Fields in : " & FindCombine
		end if
		' Remove the comments out of the Template File
		' This ensures that the Signature Default Settings are not passed on into the Signature File when they are created in the users folder
		templateHTML = RemoveComments(templateHTML)

    else
	    AddDebug "No header comments found in " & SignatureName
    END IF

    'The value is blank, but the template contains the text to update, so lets autocreate the value.
    if CheckBlankValue(templateHTML, "*|companyurl|*", companyurl) = True then 
		'Autocreate the url name 
		companyurl = lcase(SignatureName) & ".com.au"
		AddDebug "Set Variable: companyurl 'Was BLANK' -> " & companyurl
    end if


	' Set pobox to have the words PO BOX if its currently just a number
	if AutoPOBox = True then
		if IsNumeric(pobox) and len(pobox) > 0 and len(pobox) < 6 then 
			pobox = "PO Box " + pobox
			AddDebug "Set Variable: pobox 'Missing PO BOX' -> " & pobox
		end if 
	end if
	
	' Transform some values for a few variables (makes them look better)
	name = TransformText(TransformName, "name")
	title = TransformText(TransformTitle, "title")
	state = TransformText(TransformState, "state")
	companyphone = TransformText(TransformCompanyPhone, "companyphone")
	phone = TransformText(TransformPhone, "phone")
	mobile = TransformText(TransformMobile, "mobile")
	email = TransformText(TransformEmail, "email")
	city = TransformText(TransformCity, "city")
	country= TransformText(TransformCountry, "country")
	
	' Find some more values to use 
	footertext = ReadDefaultSettings("FooterText", headerHTML)
	if instr(footertext,"^") and instr(footerText,"<p>") = 0 then 
		footertext = "<p>" + footertext +"</p>"
		footertext = Replace(footertext,"^", "</p><p>")
	end if
	
	if len(footerText) > 0 then 
		footertext = Replace(footertext,"*|xmasstartyear|*",cstr(XmasStartYear))
		footertext = Replace(footertext,"*|xmasfinishyear|*",cstr(XmasFinishYear))
		footertext = Replace(footertext,"{xmasstartyear}",cstr(XmasStartYear))
		footertext = Replace(footertext,"{xmasfinishyear}",cstr(XmasFinishYear))
	end if
	
	' Create the TEXT version of the Signature
	dim MaxL
	MaxL = 0
	if len(name) > MaxL then MaxL = len(name)
	if len(title) > MaxL then MaxL = len(title)
	if len(company) > MaxL then MaxL = len(company)
	if len(name) then TemplateTEXT = TemplateTEXT + ucase(name) + VbCrLf
	if len(title) then TemplateTEXT = TemplateTEXT + title + VbCrLf
	if len(company) then TemplateTEXT = TemplateTEXT + string(MaxL+5,"-") + VbCrLf+ company + VbCrLf+ string(MaxL+5,"-") + VbCrLf
	if len(phone) then TemplateTEXT = TemplateTEXT + "D: " & phone + VbCrLf
	if len(companyphone) then TemplateTEXT = TemplateTEXT + "P: " & companyphone + VbCrLf
	if len(mobile) then TemplateTEXT = TemplateTEXT + "M: " + mobile  + VbCrLf
	if len(companyfax) then TemplateTEXT = TemplateTEXT + "F: " + companyfax + VbCrLf
	if len(address) then TemplateTEXT = TemplateTEXT + Replace(address,"^",", ") 
	if len(suburb) and instr(addresss, suburb) = 0 then TemplateTEXT = TemplateTEXT + ", " + suburb
	if len(city) and instr(address,city) = 0 and city <> suburb then TemplateTEXT = TemplateTEXT + ", "+ city
	if len(state) = 0 then TemplateTEXT = TemplateTEXT + VbCrLf
	if len(state) and state <> city then TemplateTEXT = TemplateTEXT + ", " + state +" " 
	if len(postcode) = 0 then TemplateTEXT = TemplateTEXT + VbCrLf
	if len(postcode) then TemplateTEXT = TemplateTEXT + " " + postcode + VbCrLf
	if len(pobox) then TemplateTEXT = TemplateTEXT + pobox
	if len(suburb) then TemplateTEXT = TemplateTEXT + ", " + suburb 
	if len(city) and city <> suburb then TemplateTEXT = TemplateTEXT + ", "+city
	TemplateTEXT = TemplateTEXT + VbCrLf
	if len(companyurl) then TemplateTEXT = TemplateTEXT + ucase(companyurl)
			
			
	TemplateTEXT = StripHTML(TemplateTEXT)
	
	if len(membernote) then 
	   if CheckExtra(membernote) = True then 
	   AddDebug "Set Variable: membernote 'User in Extra Group' -> " & extranote
	   end if
	end if
	
	if HideTableRows = False then 
		' Set the variable to hide via HTML if they are blank
		company_v = IIF(company = "", DisplayNone, "")
		companyurl_v = IIF(companyurl = "", DisplayNone, "")
		companyfax_v = IIF(companyfax = "", DisplayNone, "")
		companyphone_v = IIF(companyphone = "", DisplayNone, "")
		department_v = IIF(department = "", DisplayNone, "")
		countryname_v = IIF(countryname = "", DisplayNone, "")
		webpage_v =IIF(webpage = "", DisplayNone, "")
		office_v = IIF(office = "", DisplayNone, "")
		suburb_v = IIF(suburb = "", DisplayNone, "")
		postcode_v = IIF(postcode = "", DisplayNone, "")
		country_v = IIF(country = "", DisplayNone, "")
		state_v = IIF(state = "", DisplayNone, "")
		address_v = IIF(address = "", DisplayNone, "")
		city_v = IIF(city = "", DisplayNone, "")
		pobox_v = IIF(pobox = "", DisplayNone, "")
		email_v = IIF(email = "", DisplayNone, "")
		mobile_v = IIF(mobile = "", DisplayNone, "")
		phone_v = IIF(phone = "", DisplayNone, "")
		ipphone_v = IIF(ipphone = "", DisplayNone, "")
		title_v = IIF(title = "", DisplayNone, "")
		name_v = IIF(name = "", DisplayNone, "")
		firstname_v = IIF(firstname = "", DisplayNone, "")
		lastname_v = IIF(lastname = "", DisplayNone, "")
		xmastext_v = IIF(xmastext = "", DisplayNone, "")
		xmastext_t = IIF(xmastext <> "", chr(34)+chr(34), "hidden")
		extranote_t = IIF(extranote <> "", chr(34)+chr(34), "hidden")
		IDnumber_v = IIF(IDnumber = "", DisplayNone, "")
		extranote_v = IIF(extranote = "", DisplayNone, "")
		notes_v	= IIF(notes = "", DisplayNone, "")
		footertext_v = IIF(footertext = "", DisplayNone, "")
	else
		' Hiding Tables using these vales if blank
		xmastext_t = IIF(xmastext = "", "hidden", chr(34)+chr(34))
		extranote_t = IIF(extranote = "", "hidden", chr(34)+chr(34))
		xmastext_v = IIF(xmastext = "", "hidden", chr(34)+chr(34))
		extranote_v = IIF(extranote = "", "hidden", chr(34)+chr(34))
	end if
	if DisplaySocialIcons = False then 
		socialicons_v = "hidden"
	else	
		socialicons_v = chr(34) + chr(34)
	end if
	
	Dim ImageSwap, FindSwap, WithSwap
	IF MaxXmasImageSwap > 0 Then
		AddDebug "Modifying " + cstr(MaxXmasImageSwap)+" Image Names"
		Do
			ImageSwap = split(XmasImageSwap(MaxXmasImageSwap),",")
			if ubound(ImageSwap) > 0 then 
				FindSwap = ImageSwap(0)
				WithSwap = ImageSwap(1)
				if len(FindSwap) > 0 and len(WithSwap) > 0 then 
					AddDebug("Mofifying Image " + FindSwap+" -> " + WithSwap)
					templateHTML = Replace(templateHTML,FindSwap, WithSwap,1,-1,1)
				end if
			end if
			
			MaxXmasImageSwap = MaxXmasImageSwap -1
		LOOP until MaxXmasImageSwap <= 0
		
	END IF
	
	' Modify values with HTML code
	IF MaxInsertHTML > 0 THEN
		AddDebug "Modifying " + cstr(MaxInsertHTML)+" Field Values with HTML Code"

		DO	
			tmpInsertHTML = ""
			CheckAdditionalSpace = ""
			select case LCASE(InsertHTML(MaxInsertHTML))
					case "address"
					tmpInsertHTML = address
					case "notes"
					tmpInsertHTML = notes
			end select
			if len(tmpInsertHTML) then 
			
				if right(tmpInsertHTML, 1) = InsertHTMLFind(MaxInsertHTML) then 
					tmpInsertHTML = left(tmpInsertHTML,len(tmpInsertHTML)-1)
				end if
				
				MaxL = 1
				do
					FindInsertHTML = instr(MaxL,tmpInsertHTML, InsertHTMLFind(MaxInsertHTML))
					if FindInsertHTML > 1 and FindInsertHTML< len(tmpInsertHTML) then  
						tmpInsertHTML = left(tmpInsertHTML,FindInsertHTML-1) + InsertHTMLCode(MaxInsertHTML) + mid(tmpInsertHTML, FindInsertHTML+len(InsertHTMLFind(MaxInsertHTML)),len(tmpInsertHTML))
					end if
					if FindInsertHTML = 1 then
						tmpInsertHTML = InsertHTMLCode(MaxInsertHTML) + mid(tmpInsertHTML, FindInsertHTML+len(InsertHTMLFind(MaxInsertHTML)),len(tmpInsertHTML))
					end if
					if FindInsertHTML = len(tmpInsertHTML) then
						tmpInsertHTML = left(tmpInsertHTML,FindInsertHTML) + InsertHTMLCode(MaxInsertHTML)
					end if
					MaxL = FindInsertHTML + 1
				loop until FindInsertHTML = 0 or MaxL >= len(tmpInsertHTML)
				
			end if	
			
			select case LCASE(InsertHTML(MaxInsertHTML))
					case "address"
						address = tmpInsertHTML
						address = replace(address,chr(13),"",1,-1,1)
					case "notes"
						notes = tmpInsertHTML
			end select
			IF MaxL > 1 then AddDebug "Mod Variable: " & InsertHTML(MaxInsertHTML)&" 'InsertHTML' -> " & tmpInsertHTML
			MaxInsertHTML = MaxInsertHTML -1
			
		loop until MaxInsertHTML <= 0
		
	END IF
	
	' Add a HTML Space to the phone details
	if AddPhoneSpace = True then 
		if len(phone) then phone = phone + Spc
		if len(mobile) then mobile = mobile + Spc
		if len(companyphone) then companyphone = companyphone + Spc
		if len(ipphone) then ipphone = ipphone + Spc
		if len(companyfax) then companyfax = companyfax + Spc
		if len(companyphone) then companyphone = companyphone + Spc
	end if
	
	' Add a HTML Space to the address details
	IF AddAddressSpace = True then 
		if len(address) then address = address + Spc
		if len(country) then country = country + Spc
		if len(suburb) then suburb = suburb + Spc
		if len(city) then city = city + Spc
		if len(pobox) then pobox = pobox + Spc
		if len(state) then state = state + Spc
	END IF
	
	' Add an Extra HTML Spaces to some address details
	IF AddAdditionalAddressSpace = True then 
		if len(address) then address = address + Spc
	END IF
	IF AddAdditionalCitySpace = True then 
		if len(city) then city = city + Spc
	END IF
	IF AddAdditionalStateSpace = True then 
		if len(state) then state = state + Spc
	END IF
	IF AddAdditionalSuburbSpace = True then 
		if len(suburb) then suburb = suburb + Spc
	END IF
	IF AddAdditionalPOBoxSpace = True then 
		if len(pobox) then pobox = pobox + Spc
	END IF
	IF AddAdditionalTitleSpace = True then 
		if len(title) then title = title + Spc
	END IF
	IF AddAdditionalPrevTitleSpace = True then 
		if len(title) then title = Spc + title
	END IF
	
	' Combine Field Data
	IF MaxCombine > 0 then 
				
		dim ci, chi, FindField, FindCode, FindF, Minus, Plus, CombineData, UseField, LenField, CHtml,CField, FCombineLast, FCombineBlank, Merge
		dim cname, cvalue
		ci = 1
		do
			
			SplitCombine = split(Combine(ci),"+")
			CombineData = ""
			UseField = SplitCombine(0)
			chi = 1
			FCombineLast = True
			FCombineBlank = False
			if lcase(CombineLastCode(ci)) = "false" then FCombineLast = False
			if lcase(CombineBlankCode(ci)) = "true" then FCombineBlank = True
			cname = ""
			do
				CHtml = CombineHTML(ci, chi)
				CField = CombineField(ci, chi)
				
				if Chtml = "''" then CHtml = ""
				
				if len(CHtml) > 2 then 
					if left(CHtml,1) = "'" then CHtml = mid(CHtml, 2, len(CHtml))
					if right(CHtml,1) = "'" then CHtml = left(CHtml, len(CHtml)-1)
				end if
				
				FindField = CField
				FindCode = CHtml
				LenField = len(CHtml)
					
				'AddDebug UseField &" " & chi&" " & FindField &" " &FindCode
				Merge = False:Minus = False:Plus = False
				if left(CField, 1) = "-" then 
					Minus = True
					CField = mid(CField, 2, Len(CField))
				end if	
				if left(CField, 1) = "+" then 
					Plus = True
					CField = mid(CField, 2, Len(CField))
				end if
				if left(CField, 1) = "#" then 
					Merge = True
					CField = mid(CField, 2, Len(CField))
				end if
				if len(CField) and CField <> "''" then 
					if len(Eval(CField)) > 0 or FCombineBlank = True then 
						if chi = MaxCombineHTML(ci) and FCombineLast = False then FindCode = ""
						
						if FindCode <> "" then 
							FindCode  = Replace(FindCode ,"href=" + chr(34) + "tel:*|"&CField&"|*","href=" + chr(34) + "tel:"+replace(StripHTML(Eval(CField))," ", ""), 1, -1, 1)
							FindCode  = Replace(FindCode ,"href =" + chr(34) + "tel:*|"&CField&"|*","href=" + chr(34) + "tel:"+replace(StripHTML(Eval(CField))," ", ""), 1, -1, 1)
							FindCode  = Replace(FindCode ,"href= " + chr(34) + "tel:*|"&CField&"|*","href=" + chr(34) + "tel:"+replace(StripHTML(Eval(CField))," ", ""), 1, -1, 1)
							FindCode  = Replace(FindCode ,"href = " + chr(34) + "tel:*|"&CField&"|*","href=" + chr(34) + "tel:"+replace(StripHTML(Eval(CField))," ", ""), 1, -1, 1)
							
							FindCode  = Replace(FindCode ,"href=" + chr(34) + "mailto:|*"&CField&"|*","href=" + chr(34) + "mailto:" + replace(StripHTML(Eval(CField))," ", ""), 1, -1, 1)
							FindCode  = Replace(FindCode ,"href =" + chr(34) + "mailto:*|"&CField&"|*","href=" + chr(34) + "mailto:" + replace(StripHTML(Eval(CField))," ", ""), 1, -1, 1)
							FindCode  = Replace(FindCode ,"href= " + chr(34) + "mailto:*|"&CField&"|*","href=" + chr(34) + "mailto:" + replace(StripHTML(Eval(CField))," ", ""), 1, -1, 1)
							FindCode  = Replace(FindCode ,"href = " + chr(34) + "mailto:*|"&CField&"|*","href=" + chr(34) + "mailto:" + replace(StripHTML(Eval(CField))," ", ""), 1, -1, 1)
							
							FindCode  = Replace(FindCode ,"*|"&CField&"|*", StripHTML(Eval(CField)), 1, -1, 1)
							FindCode  = Replace(FindCode ,"*|HighlightColor|*", HighlightColor, 1, -1, 1) 
							FindCode  = Replace(FindCode ,"*|TextColor|*", TextColor, 1, -1, 1) 
							FindCode  = Replace(FindCode ,"*|BarColor|*", BarColor, 1, -1, 1) 
							FindCode  = Replace(FindCode ,"*|TextColorHighlight|*", TextColorHighlight,1,-1,1) 
						
							if Plus = True then CombineData = CombineData + FindCode + Eval(CField)
							if Minus = True then CombineData = CombineData + Eval(CField) + FindCode
							if merge = True then CombineData = CombineData + FindCode
						else
							if Minus = False and Plus = False and Merge = False then CombineData = CombineData + Eval(CField)
						end if
					end if
				end if	
				if chi = 2 then cname = CField
				if chi > 2 then cname = cname + "+" + CField
				chi = chi + 1
				
			loop until chi > MaxCombineHTML(ci)
			
			select case lcase(UseField)
				case "name"
					name = CombineData
				case "title"
					title = CombineData
				case "address"
					address = CombineData
				case "suburb"
					suburb = CombinedData
				case "state"
					state = CombinedData
				case "pobox"
					pobox = CombineData
				case "companyphone"
					companyphone = CombineData
				case "mobile"
					mobile = CombineData
				case "phone"
					phone = CombineData
				case "companyfax"
					companyfax = CombineData
				case "webpage"
					webpage = CombineData
			end select
			
			AddDebug "Mod Variable: "& UseField&" 'Combine " &cname& "' -> " & CombineData
			ci = ci + 1
			
		loop until ci > MaxCombine
	END IF
	
    ' Make global template value replacements in the HTML code
	templateHTML = Replace(templateHTML, "*|name|*", name, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|name_v|*", name_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|title|*", title, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|title_v|*", title_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "href=" + chr(34) + "tel:*|phone|*","href=" + chr(34) + "tel:" + replace(StripHTML(phone)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href =" + chr(34) + "tel:*|phone|*","href=" + chr(34) + "tel:" + replace(StripHTML(phone)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href= " + chr(34) + "tel:*|phone|*","href=" + chr(34) + "tel:" + replace(StripHTML(phone)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href = " + chr(34) + "tel:*|phone|*","href=" + chr(34) + "tel:" + replace(StripHTML(phone)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "*|phone|*", phone,1,-1,1)
	templateHTML = Replace(templateHTML, "*|phone_v|*", phone_v,1,-1,1)
	templateHTML = Replace(templateHTML, "*|ipphone|*", ipphone,1,-1,1)
	templateHTML = Replace(templateHTML, "*|ipphone_v|*", ipphone_v,1,-1,1)
	templateHTML = Replace(templateHTML, "href=" + chr(34) + "tel:*|mobile|*","href=" + chr(34) + "tel:"+replace(StripHTML(mobile)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href =" + chr(34) + "tel:*|mobile|*","href=" + chr(34) + "tel:"+replace(StripHTML(mobile)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href= " + chr(34) + "tel:*|mobile|*","href=" + chr(34) + "tel:"+replace(StripHTML(mobile)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href = " + chr(34) + "tel:*|mobile|*","href=" + chr(34) + "tel:"+replace(StripHTML(mobile)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "*|mobile|*", mobile,1,-1,1)
	templateHTML = Replace(templateHTML, "*|mobile_v|*", mobile_v,1,-1,1)
	templateHTML = Replace(templateHTML, "href=" + chr(34) + "mailto:*|email|*","href=" + chr(34) + "mailto:" + replace(StripHTML(email)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href =" + chr(34) + "mailto:*|email|*","href=" + chr(34) + "mailto:" + replace(StripHTML(email)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href= " + chr(34) + "mailto:*|email|*","href=" + chr(34) + "mailto:" + replace(StripHTML(email)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href = " + chr(34) + "mailto:*|email|*","href=" + chr(34) + "mailto:" + replace(StripHTML(email)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "*|email|*", email,1,-1,1)
	templateHTML = Replace(templateHTML, "*|email_v|*", email_v,1,-1,1)

	templateHTML = Replace(templateHTML, "*|highlightcolor|*", HighlightColor, 1, -1, 1)      
	templateHTML = Replace(templateHTML, "*|signaturename|*", SignatureName, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|dayname|*", cdayname, 1, -1, 1)
    templateHTML = Replace(templateHTML, "*|monthname|*", cmonthname, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|day|*",cstr(cday), 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|dday|*",cdday, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|month|*",cstr(cmonth), 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|mmonth|*",cmmonth, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|year|*",cstr(cyear), 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|nextyear|*",cstr(cyear+1), 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|prevyear|*",cstr(cyear-1), 1, -1, 1)
		

  	templateHTML = Replace(templateHTML, "*|company|*",company, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|company_v|*",company_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "href=" + chr(34) + "tel:*|companyphone|*","href=" + chr(34) + "tel:"+replace(StripHTML(companyphone)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href =" + chr(34) + "tel:*|companyphone|*","href=" + chr(34) + "tel:"+replace(StripHTML(companyphone)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href= " + chr(34) + "tel:*|companyphone|*","href=" + chr(34) + "tel:"+replace(StripHTML(companyphone)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href = " + chr(34) + "tel:*|companyphone|*","href=" + chr(34) + "tel:" + replace(StripHTML(companyphone)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "*|companyphone|*",companyphone, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|companyphone_v|*",companyphone_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|companyfax|*",companyfax, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|companyfax_v|*",companyfax_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "href=" + chr(34) + "*|companyurl|*","href=" + chr(34) + replace(StripHTML(companyurl)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href =" + chr(34) + "*|companyurl|*","href=" + chr(34) + replace(StripHTML(companyurl)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href= " + chr(34) + "*|companyurl|*","href=" + chr(34) + replace(StripHTML(companyurl)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "href = " + chr(34) + "*|companyurl|*","href=" + chr(34) + replace(StripHTML(companyurl)," ",""),1,-1,1)
	templateHTML = Replace(templateHTML, "*|companyurl|*",companyurl, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|companyurl_v|*",companyurl_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|address|*",address, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|address_v|*",address_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|city|*",city, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|city_v|*",city_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|state|*",state, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|state_v|*",state_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|country|*",country, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|country_v|*",country_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|postcode|*",postcode, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|postcode_v|*",postcode_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|pobox|*",pobox, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|pobox_v|*",pobox_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|suburb|*",suburb, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|suburb_v|*",suburb_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|firstname|*",firstname, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|firstname_v|*",firstname_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|lastname|*",lastname, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|lastname_v|*",lastname_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|xmastext|*",xmastext, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|xmastext_v|*",xmastext_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|xmastext_t|*",xmastext_t, 1, -1, 1)
	templateHTML = Replace(templateHTML, "href=" + chr(34) + "*|webpage|*","href=" + chr(34) + StripHTML(webpage),1,-1,1)
	templateHTML = Replace(templateHTML, "href =" + chr(34) + "*|webpage|*","href=" + chr(34) + StripHTML(webpage),1,-1,1)
	templateHTML = Replace(templateHTML, "href= " + chr(34) + "*|webpage|*","href=" + chr(34) + StripHTML(webpage),1,-1,1)
	templateHTML = Replace(templateHTML, "href = " + chr(34) + "*|webpage|*","href=" + chr(34) + StripHTML(webpage),1,-1,1)
	templateHTML = Replace(templateHTML, "*|webpage|*",webpage, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|webpage_v|*",webpage_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "href=" + chr(34) + "*|href|*","href=" + chr(34) + StripHTML(href),1,-1,1)
	templateHTML = Replace(templateHTML, "href =" + chr(34) + "*|href|*","href=" + chr(34) + StripHTML(href),1,-1,1)
	templateHTML = Replace(templateHTML, "href= " + chr(34) + "*|href|*","href=" + chr(34) + StripHTML(href),1,-1,1)
	templateHTML = Replace(templateHTML, "href = " + chr(34) + "*|href|*","href=" + chr(34) + StripHTML(href),1,-1,1)
    templateHTML = Replace(templateHTML, "*|IDnumber|*",IDnumber, 1, -1, 1)	
	templateHTML = Replace(templateHTML, "*|IDnumber_v|*",IDnumber_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|extranote|*",extranote,1 , -1 , 1)
	templateHTML = Replace(templateHTML, "*|extranote_v|*",extranote_v, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|extranote_t|*",extranote_t, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|notes|*",notes, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|notes_v|*",notes_v,1,-1,1)
	templateHTML = Replace(templateHTML, "*|footertext|*",footertext, 1, -1, 1)
	templateHTML = Replace(templateHTML, "*|footertext_v|*",footertext_v, 1, -1 ,1)
	templateHTML = Replace(templateHTML, "*|socialicons_v|*",socialicons_v, 1, -1, 1)

	
	if len(XmasOverWriteColor) > 0 and len(xmashighlight) > 0 then 
	   
	   if instr(lcase(XmasOverWriteColor),"linkwwwcolor") > 0 then templateHTML = Replace(templateHTML,"*|LinkWWWColor|*",xmashighlight,1,-1,1)
	   if instr(lcase(XmasOverWriteColor),"textcolorhighlight") > 0 then templateHTML = Replace(templateHTML,"*|TextColorHighlight|*",xmashighlight,1,-1,1)
	   
	   AddDebug "Overwite with Xmas Color: " & XmasOverWriteColor &" = " & xmashighlight
	end if
	templateHTML = Replace(templateHTML,"*|LinkWWWColor|*",LinkWWWColor,1,-1,1)
	templateHTML = Replace(templateHTML,"*|SymbolColor|*",SymbolColor,1,-1,1)
	
	templateHTML = Replace(templateHTML,"*|TextColor|*",TextColor,1,-1,1)
	
	templateHTML = Replace(templateHTML,"*|TextColorHighlight|*",TextColorHighlight,1,-1,1)
	templateHTML = Replace(templateHTML,"*|TextFooterColor|*",TextFooterColor,1,-1,1)
	templateHTML = Replace(templateHTML,"*|BarColor|*",BarColor,1,-1,1)
	templateHTML = Replace(templateHTML,"*|HyperlinkColor|*",HyperlinkColor,1,-1,1)
	templateHTML = Replace(templateHTML,"*|DefaultFont|*",DefaultFont,1,-1,1)
	templateHTML = Replace(templateHTML,"*|DefaultLineHeight|*",DefaultLineHeight,1,-1,1)
	templateHTML = Replace(templateHTML,"*|DefaultFontSize|*",DefaultFontSize,1,-1,1)
	templateHTML = Replace(templateHTML,"*|LargerFontSize|*",LargerFontSize,1,-1,1)
	
	' If using the imagename variable, then update the HTML code
	if len(imagename) then 
		templateHTML = Replace(templateHTML,"*|imagename|*",imagename,1,-1,1)
	else
		' If the template is using the imagename variable but not defined then make it up
		AddDebug "Set Variable: imagename 'Auto' -> " & SignatureName&"."& DefaultImageType
        templateHTML = Replace(templateHTML,"*|imagename|*",SignatureName&"." & DefaultImageType,1,-1,1)
	end if

    ' Modify the Width and Height in the HTML code
	templateHTML = Replace(templateHTML,"*|width|*",SignatureWidth,1,-1,1)
    templateHTML = Replace(templateHTML,"*|height|*",SignatureHeight,1,-1,1)

	IF MaxSocialIcons > 0 then 
		Socialtmpname=SignatureName
		' Set the Social Icon Name from the Signatures Name
		if len(SocialIconName) > 0 then Socialtmpname = SocialIconName
		' Use another signatures social icon pictures Instead
		if len(ForceSocialIcon) > 0 then Socialtmpname = ForceSocialIcon
		dim si
		si = 1
		do
			templateHTML = Replace(templateHTML,"*|" + SocialIcon(si) + "|*",Socialtmpname & "-" & SocialIcon(si) & "." & DefaultImageType,1,-1,1)		
			templateHTML = Replace(templateHTML,"*|" + SocialIcon(si) + "Link|*",SocialIconLink(si),1,-1,1)		
			templateHTML = Replace(templateHTML,"*|" + SocialIcon(si) + "Alt|*",SocialIconAlt(si),1,-1,1)		
			si = si + 1
		loop until si > MaxSocialIcons
	END IF
	
	

	'SetVariables templateHTML
	
	' Display Active Directory values in the debug. makes it easier to check debug file, rather than checking 2 files at once.
	AddDebug ""
	if len(name) then AddDebug         "name         :" & name
	if len(firstname) then AddDebug    "firstname    :" & firstname
	if len(lastname) then AddDebug     "lastname     :" & lastname
	if len(title) then AddDebug        "title        :" & title
	if len(email) then AddDebug        "email        :" & email
	if len(companyphone) then AddDebug "companyphone :" & companyphone
	if len(mobile) then AddDebug       "mobile       :" & mobile
	if len(phone) then AddDebug        "phone        :" & phone
	if len(ipphone) then AddDebug      "ipphone      :" & ipphone
	if len(address) then AddDebug      "address      :" & address
	if len(suburb) then AddDebug       "suburb       :" & suburb
	if len(city) then AddDebug         "city         :" & city
	if len(state) then AddDebug        "state        :" & state
	if len(postcode) then AddDebug     "postcode     :" & postcode
	if len(pobox) then AddDebug        "pobox        :" & pobox
	if len(country) then AddDebug      "country      :" & country
	if len(imagename) then AddDebug    "imagename    :" & imagename
	if len(extranote) then AddDebug    "extranote    :" & extranote
	if len(notes) then AddDebug   	   "notes        :" & notes
	if len(webpage) then AddDebug 	   "webpage      :" & webpage
	if len(whenchanged) then AddDebug  "modified     :" & whenchanged
	if len(xmastext) then AddDebug     "xmas         :" & xmastext
	if len(footertext) then AddDebug   "footer       :" & footertext
	if len(IDnumber) then AddDebug     "IDnumber     :" & IDnumber
	AddDebug ""
	
	' Create the Users HTML signature file
	Dim htmlSignatureFile, textSignatureFile
	on error resume next
	
	Set htmlSignatureFile = fileSystem.CreateTextFile(signaturesFolderPath & SignatureName & ".htm", True)
	htmlSignatureFile.Write(templateHTML)	
	htmlSignatureFile.Close
	Set htmlSignatureFile = Nothing
	
	' Create the Users TEXT Signature file
	Set textSignatureFile = fileSystem.CreateTextFile(signaturesFolderPath & SignatureName & ".txt", True)
	textSignatureFile.Write(templateTEXT)	
	textSignatureFile.Close
	Set textSignatureFile = Nothing
	on error goto 0
	
		
    ' Download the required image files
    IF len(trim(imagename)) > 0 THEN
	    DownloadFile imagename, sourceFilesUrl, signaturesFolderPath &SignatureImageFolder
    ELSE
         ' The imagename value was not set in the Signature Template file, so find all the imagefiles in the html code and download them.
         DownloadAllFiles templateHTML,sourceFilesUrl, signaturesFolderPath &SignatureImageFolder
    END IF
    
	' Download an additional imagename.
	IF len(aditionalimage) > 0 THEN 
		DownloadFile aditionalimage, sourceFilesUrl, signaturesFolderPath &SignatureImageFolder
	END IF
	
	' Download any SocialIcons Images
	IF MaxSocialIcons > 0 then 
		DO
			Socialtmpname=SignatureName
			IF LEN(SocialIconName) > 0 THEN Socialtmpname=SocialIconName
			imagename = Socialtmpname +"-"+SocialIcon(MaxSocialIcons)
			IF INSTR(imagename,".") = 0 then 
				IF INSTR(imagename,"." + DefaultImageType) = 0 then imagename=imagename +"." + DefaultImageType	
			END IF
			DownloadFile imagename, sourceFilesUrl, signaturesFolderPath &SignatureImageFolder
			MaxSocialIcons = MaxSocialIcons - 1
		loop until MaxSocialIcons = 0
	END IF
	
	' Download any Extra Files
	IF MaxExtraFile > 0 then 
		Do
		  imagename = ExtraFile(MaxExtraFile)
		  IF INSTR(imagename,".") = 0 then 
			IF INSTR(imagename,"." + DefaultImageType) = 0 then imagename = imagename +"." + DefaultImageType	
		  END IF
		  DownloadFile imagename, sourceFilesUrl, signaturesFolderPath &SignatureImageFolder
		  MaxExtraFile = MaxExtraFile - 1
		loop until MaxExtraFile = 0
	END IF
	
	' Create the VCF Signature contact card 
	CreateVCard signaturesFolderPath , SignatureName
	
	' Signature Complete
	AddDebug "Create Signature Completed."
	AddDebug ""
	shell.LogEvent 0, userName & " - Created Outlook Signature " + SignatureName

END IF
	
END SUB

Function ReplaceEmailDomain (strEmail, strReplace)

	DIM ReplaceEmail, a, strFind, strReplacement, strStart
   
	ReplaceEmail = split(strReplace,",")
	if ubound(ReplaceEmail) > 1 then 
		strFind = ReplaceEmail(0)
		strStart = strEmail
		strReplacement = ""
		a = 1
		do
		    strReplacement = strReplacement+ReplaceEmail(a) + " "
			strEmail = Replace(strEmail, ReplaceEmail(a), strFind)
			a = a + 1
		loop until a > ubound(ReplaceEmail)
		'AddDebug "Replacing the domain name in email address "& strStart &" with " & strFind &" if found in these " & strReplacement
		if lcase(strEmail) <> lcase(strStart) then 
			AddDebug "Mod Variable: email 'EmailDomain' -> " & strEmail
		end if
	end if
   ReplaceEmailDomain = strEmail
   
End Function

Function SetToDefault(strVarname, strdefault)
	
	Dim varresult
	varresult = Eval(strVarname)
	
	if eval(strdefault) = chr(34) + chr(34) then 
		SetToDefault = ""
		AddDebug "Set Variable: "+strVarname& " '" &strdefault&"' -> BLANK"
	else
        if len(eval(strdefault)) = 0 then 
			SetToDefault=varresult
		else
			SetToDefault=eval(strdefault)
			AddDebug "Set Variable: " &strVarname& " '" &strdefault& "' -> " &SetToDefault
		end if
	end if
end function

Function RemoveItemFromHTML(byval templateHTML, byval Start, byval Finish)
	
	if instr(templateHTML, Start) > 0 and instr(templateHTML, Finish) > 0 THEN
		Dim p1, p2, part1
		p1 = instr(templateHTML, Start)
		if p1 > 0 then 
			p2 = instr(p1+len(Start),templateHTML,Finish)
			if p2 > 0 then 
				if p1 = 1 then p1 = 2
				part1 = left(templateHTML, p1-1) + mid(templateHTML, p2 + len(Finish), len(templateHTML))
				if len(part1) > 0 then 
					RemoveItemFromHTML = part1
					AddDebug "Removing HTML Code Between '" & Start&"' and '" & Finish &"'"
				end if
			end if
		end if	
	Else
		AddDebug "Did Not Remove HTML Code Between '" & Start&"' and '" & Finish &"'"
		RemoveItemFromHTML = templateHTML
	END IF
END Function

Function ReadHeaderSignatureHTML(byval templateHTML, byval SignatureName)

	Dim p1, p2, part1
    p1 = instr(templateHTML,"<!--")
	
	' The first comment in the templateHTML is used for signature settings.
	' It must also have the signaturename mentioned as it confirms that it is a header file
    if p1 then 
            p2 = instr(p1+5,templateHTML,"-->")
            if p2 then 
				part1 = mid(templateHTML, p1, p2-p1)
				' If the Headerfile is not a Signature Header File then return nothing
				if instr(lcase(part1), lcase(SignatureName & ".tpl")) <> 0 then     
					ReadHeaderSignatureHTML = part1
					exit function
		        else
				AddDebug "Header not used as " & SignatureName &".tpl not found in header comments. - FAILED"
				end if
			end if
		end if	
	
	ReadHeaderSignatureHTML = ""
		
End Function

Function CheckBlankValue(byval templateHTML, CheckValue, FindValue)
   
       if instr(templateHTML,CheckValue) then 
          if len(trim(FindValue)) = 0 then 
              CheckBlankValue=true
          else
             CheckBlankValue=False
          end if
       else
          CheckBlankValue=False
       end if

End Function

Function RemoveLastName(byval email)
	
	' Remove the lastname from the email address
	Dim p1, email1, email2, p2
	RemoveLastName = ""

	p1 = instr(email,"@")
	if p1 > 1 then 
		email1 = left(email, p1-1)
		email2 = mid(email, p1, len(email))
		p2 = instr(email1, ".")
		if p2 > 1 then 
			email1 = left(email1, p2-1)
		end if
		RemoveLastName = email1 + email2
	else
		RemoveLastName = email
	end if

End function


Function ChangeDomain(byval email, newdomainname)

	' Update the Users Domain in their email address
	Dim p1
	ChangeDomain = ""

	' Add an @ to the newdomainname if it does not have one.
	if left(newdomainname,1) <> "@" then newdomainname = "@" + newdomainname

	' Remove the lastname from the email address
	p1 = instr(email, "@")
	if p1 > 1 and len(newdomainname) > 1 then       
	    ChangeDomain = left(email, p1-1)+newdomainname
	else
	    ChangeDomain = email + newdomainname
	end if

end function

Function RemoveComments (byval templateHTML)

	' Strip out the first comment so that the users signature does not contain these settings in the new html signature
	Dim p1,p2, part1
    p1 = instr(templateHTML, "<!--")
    
    if p1 then 
		part1 = left(templateHTML, p1-1)
        p2 = instr(p1+4,templateHTML, "-->")
        if p2 < len(templateHTML)-2 then 
            RemoveComments = part1 + mid(templateHTML, p2 + 3, len(templateHTML))
			AddDebug "headerHTML settings stripped from Signature File."
            exit function
        end if
        RemoveComments = templateHTML
    end if

End Function

Function ReadDefaultSettings(CheckDefault, headerHTML) 
	
	' Find a DefaultSetting in the Header File
    ReadDefaultSettings = ""
    Dim a, b, c, d, strText, strBlank, strQuote, fshtml, fehtml
    d = 4 + len(CheckDefault)
    strBlank = ""
    b = 1
	
	do

	a = instr(b,headerHTML,"<*|"+CheckDefault+"|*")
	
    IF a then 
		strQuote= mid(headerHTML, a + d + 1,1)
		'if the value starts with a ' then search for the closing end as well
		if strQuote = "'" then 
	        if mid(headerHTML, a + d + 1, 3) ="''>" then 
				c = a + d + 3
			else
				c = instr(a + d + 1, headerHTML, "'>") ' the next best end of line to fine
				if c = 0 then c = instr(a + d + 1, headerHTML, "' >") ' the last end of line to find!
				if c <> 0 then c = c + 1
			end if
		else
			strQuote = ""
			c = instr(a + d, headerHTML, ">")
			if c = 0 then c = instr(a + d, headerHTML, VbCrlf)
		end if
	   
		b = c + 1
		' find the closing '>' character 
		IF c-a-d-1 > 1 then 
			strText = mid(headerHTML, a + d + 1, c - a - d - 1)
			if strText = chr(34) + chr(34) then strBlank = "(Blank)"

			if strText <> string(len(strText),"*") then 
				ReadDefaultSettings = strText
				if len(ReadDefaultSettings) > 3 then 
					if left(ReadDefaultSettings,1) = "'" and right(ReadDefaultSettings,1) = "'" then 
						ReadDefaultSettings = mid(ReadDefaultSettings, 2, len(ReadDefaultSettings) - 2)
					Else
						' interpret the {...} EG: HTML code and replace with the equavilent <...> EG: <DIV> </DIV> etc..
						fshtml = instr(ReadDefaultSettings,"{")
						fehtml = instr(ReadDefaultSettings,"}")
						if fshtml > 0 and fehtml > 0 then 
							Do
								fshtml = instr(ReadDefaultSettings	,"{")
								fehtml = instr(ReadDefaultSettings, "}")
								if fshtml > 0 and fehtml > 0 then 
									ReadDefaultSettings = left(ReadDefaultSettings, fshtml -1) +"<" + mid(ReadDefaultSettings, fshtml +1,len(ReadDefaultSettings))
									ReadDefaultSettings = left(ReadDefaultSettings, fehtml -1) +">" + mid(ReadDefaultSettings, fehtml +1,len(ReadDefaultSettings))
								END IF
							loop until fshtml = 0 or fehtml = 0
						END IF
					end if
				end if
				

				AddDebug "Set Variable: " + CheckDefault+" -> " + ReadDefaultSettings  & strBlank  
				' Modify HeaderHTML so the next occurance of this same checkdefault can be found
				headerHTML = left(headerHTML, a + d) + string(len(strText), "*") + mid(headerHTML, a + d + len(strText) + 1, len(headerHTML))
				a = 0
			end if
		END IF
    END IF
	loop until a = 0 

End Function

Sub AddVarDebug(varname)
	
	' Add Variable Information to the Debug File (debug.txt)
	
    Dim varresult
	varresult = Eval(varname)
	if len(varresult) then AddDebug "Set Variable: " & varname & " -> " & varresult
	
END Sub

Sub AddDebug(strText)

	' Add Information to the Debug File (debug.txt) in each users Signature Folder
	
	on error resume next

	if use_debug = true then 
		
		Set debugfile = fileSystem.openTextFile(signaturesFolderPath&"debug.txt", 8, True)    
		if len(strText) then 
			if left(strText,1) <> "*" then strText = FormatDateTime(Now, vbShortTime) + " " + strText
		END IF
		strText = replace(strText, vbCrLf, " ")
		debugfile.WriteLine(strText)
		debugfile.close
	
		Set debugfile = Nothing
	end if
	on error goto 0
	
End sub

Sub UpdateGlobalVarsFromAD() 

    ' Get the logged on users details
    Dim user, domainName, userLDAP, userDN, wshShell
	Dim p1, p2, e1, e2, findXmas, findXmasSignature, checkpos, findPhone
	Set wshShell = CreateObject( "WScript.Shell" )
    Set user = CreateObject("WScript.Network")
	
	ON ERROR RESUME NEXT
	dim objRootDSE, strDNSDomain, objTrans, strNetBIOSDomain
	' Determine DNS name of domain from RootDSE.
	Set objRootDSE = GetObject("LDAP://RootDSE")
	strDNSDomain = objRootDSE.Get("defaultNamingContext")

	' Use the NameTranslate object to find the NetBIOS domain name from the
	' DNS domain name.
	Set objTrans = CreateObject("NameTranslate")
	objTrans.Init ADS_NAME_INITTYPE_GC, ""
	objTrans.Set ADS_NAME_TYPE_1779, strDNSDomain
	strNetBIOSDomain = objTrans.Get(ADS_NAME_TYPE_NT4)
	' Remove trailing backslash.
	domainName = Left(strNetBIOSDomain, Len(strNetBIOSDomain) - 1)

    userName = user.UserName
	if len(userName) = 0 then 
		userName = wshShell.ExpandEnvironmentStrings("%USERNAME%")
	end if
	
	if len(domainName) = 0 then 
       domainName = user.UserDomain
	END IF
	IF len(domainName) = 0 then 
		domainName = wshShell.Environment("Process").Item("userdomain")
	end if
		
	' Get the users details from Active Directory
	AddDebug "Username: " & UserName &"  Domain: " & domainName
	
	UserDN = GetUserDN(userName,domainName)
    AddDebug "Geting Active Directory data from LDAP://" & UserDN

    Set userLDAP = GetObject("LDAP://" & UserDN)
	
	if userLDAP is not Nothing then 
		name = userLDAP.displayName
		title = userLDAP.title
		phone = userLDAP.telephoneNumber
		mobile = userLDAP.mobile
		email = userLDAP.mail
		address = userLDAP.streetAddress
		pobox = userLDAP.postOfficeBox
		state = userLDAP.st
		city = UserLDAP.l
		suburb = userLDAP.l
		country = userLDAP.c
		postcode = userLDAP.postalCode
		office = userLDAP.physicalDeliveryOfficeName
		webpage = userLDAP.wWWHomePage
		countryname = userLDAP.co
		department = userLDAP.department
		firstname = userLDAP.givenName
		lastname = userLDAP.sn
		ipphone= userLDAP.ipPhone
		WhenChanged = UserLDAP.WhenChanged
		notes = UserLDAP.info
	END IF
	
	ON ERROR GOTO 0
	
	' These values are saved so the original values can be re-evaulated if they get overwritten 
	ADAddress = address
	ADTitle = title
	ADCompany = company
	ADPhone = phone
	ADMobile = mobile
	ADstate = state
	ADpostcode = postcode	
	
	' userLDAP.otherMobile
	' userLDAP.otherIpPhone
	' userLDAP.homePhone
	' userLDAP.otherHomePhone
	' userLDAP.carLicense
	' userLDAP.roomNumber
	' UserLDAP.division
	' userLDAP.mailNickname
	' userLDAP.employeeID
	' userLDAP.employeeType
	
	' Clean the notes variable so nothing is processed, but the notesvariable can still have the old data
	if instr(notes,"-notes") > 0 then notes = ""
	
	' Check the notes field for special instructions.
	' This way user can overwrite any standard signature values.
	' If a '-' then blank that value so it does not showup on the signature.
	' if a '+' then modify the current value with this new value
	' or set a flag if otherwise found.'
		
	if len(notes) > 0 then 
		if instr(notes, "-mobile") then 
			mobile = ""
			ADmobile = ""
			notes = replace(notes, "-mobile", "")
		end if
		if instr(notes,"-phone") then 
			notes = replace(notes, "-phone", "")
			phone = ""
			ADphone = ""
		end if
		if instr(notes, "-ipphone") then 
			notes = replace(notes, "-ipphone", "")
			ipphone = ""
		end if
		if instr(notes, "-address") then 
			notes = replace(notes, "-address", "")
			address = ""
			ADaddress = ""
		end if
		if instr(notes, "-city") then 
			notes = replace(notes, "-city", "")
			city = ""
		end if
		if instr(notes,"-suburb") then 
			notes = replace(notes,"-suburb","")
			suburb = ""
		end if
		if instr(notes,"-state") then 
			notes = replace(notes,"-state","")
			state = ""
			ADstate = ""
		end if
		if instr(notes,"-department") then 
			notes = replace(notes,"-department","")
			department = ""
		end if
		if instr(notes,"-office") then 
			notes=replace(notes,"-office","")
			office=""
		end if
		if instr(notes,"-pobox") then 
			notes=replace(notes,"-pobox","")
			pobox=""
		end if
		if instr(notes,"-title") then 
			title=""
			ADtitle=""
		end if
		
		if instr(notes,"-lastname") then 
			'Remove the Last name from the Name as well...?
			IF instr(name, lastname) > 0 and len(name) > 0 and len(lastname) > 0 then 
				name = replace(name, lastname, "", 1, -1, 1)
			end if
			lastname = ""
		end if
		if instr(notes,"-firstname") then 
			'Remove the Firstname from the Name as well...?
			IF instr(name, firstname) > 0 and len(name) > 0 and len(firstname) > 0 then 
				name = replace(name, firstname, "", 1, -1, 1)
			end if
			firstname = ""
		end if
		if instr(notes,"+firstname(") then 
			findName = GetText(1,notes, "+name(",")","")
			if len(findName) > 0 then 
				firstname = findName
			end if
		end if
		if instr(notes,"-name") then 
			name = ""
		end if
		if instr(notes,"-social") then 
			DisplaySocialIcons = False
		end if
		if instr(notes,"-country") then 
			country = ""
		end if
		if instr(notes,"-postcode") then 
			postcode = ""
			ADpostcode = ""
		end if
		if instr(notes,"+name(") then 
			findName = GetText(1,notes, "+name(",")","")
			if len(findName) > 0 then 
				name = findName
			end if
		end if
		if instr(notes,"-companyphone") then 
			RemoveCompanyPhone = True
		end if
		if instr(notes,"+companyphone(") then 
			findPhone = GetText(1,notes, "+companyphone(",")","")
			if len(findphone) > 0 then 
				OverrideCompanyPhone = findphone
			end if
		end if
		if instr(notes,"testxmas") then 
			TestXmas = True
		end if
		if instr(notes,"-default") then 
			AdminDefaultSignature = True
		end if
		if instr(lcase(notes),"+xmasmessage(") then
			' overwrites the xmas message for this particular person
			findXmas = GetText(1,notes, "+xmasmessage(",")","")
			findXmasSignature = GetText(1,notes, "+xmassignature(",")","")
			if len(findXmas) > 0 and len(findXmasSignature) > 0 then 
				xmasoverwrite = findXmas
				xmassignature = findXmasSignature
			end if 
		end if

	end if
	
    Set User = Nothing
    Set userLDAP = Nothing
	Set wshShell = Nothing
	
End Sub

Function GetText(byval starthere, byval strString, byval strfind, byval strEnd, byval afterthis)
	
	' Return the Text where the Text is found in a String
	
	dim p0, p1, p2
	GetText = ""
	if starthere = 0 then starthere = 1
	p0 = starthere
	if len(strfind) > 0 then 
		if len(afterthis) > 0 then 
			' set the start point to be here in the string
			p0 = instr(starthere, strString, afterthis)
			' increment it past the search value
			if p0 > 0 then p0 = p0 + len(afterthis) + 1
		end if
		p1 = instr(p0, strString, strfind)
		if p1 > 0 then 
			p2 = instr(p1+len(strfind),strString, strEnd)
			if p2 > 0 then 
				GetText = mid(strString, p1 + len(strfind), p2 - p1 - len(strfind))
			end if
		end if
	end if
	
End function

Function GetTextPos(byval starthere, byval strString, byval strfind, byval strEnd, byval afterthis)
	
	' Return the Character Position (number) where the Text is found in a String
	
	dim p0, p1, p2
	GetTextPos = 0
	if starthere = 0 then starthere = 1
	p0 = starthere
	if len(strfind) > 0 then 
		if len(afterthis) > 0 then 
			' set the start point to be here in the string
			p0 = instr(starthere, strString, afterthis)
			' increment it past the search value
			if p0 > 0 then p0 = p0 + len(afterthis) + 1
		end if
		p1 = instr(p0, strString, strfind)
		if p1 > 0 then 
			p2 = instr(p1+len(strfind),strString, strEnd)
			if p2 > 0 then 
				GetTextPos = p1
			end if
		end if
	end if
	
End function

Function GetOutlookSignatureHtml(byval templateFilePath)

	' Read the Users Signature File
    GetOutlookSignatureHtml =""
    on error resume next

    if len(templateFilePath) > 0 then     
		Dim signatureTemplateHTMLStream
		on error resume next
			Set signatureTemplateHTMLStream = fileSystem.OpenTextFile(templateFilePath, 1, False)        
			Dim signatureTemplateHTML: signatureTemplateHTML = signatureTemplateHTMLStream.ReadAll         
		on error goto 0
		Set signatureTemplateHTMLStream = Nothing
    end if

    on error goto 0
    GetOutlookSignatureHtml = signatureTemplateHTML
    
End Function


SUB DownloadAllFiles(byval templateHTML,sourceUrl, destinationDirectory)

	' Download All the Image Files found in the Template
	AddDebug "Downloading all <img src = Files "

    Dim a,b,c,d, Part1, Part2, Part3, Part4, filename,OK
    Part1 = "<img "
	' Seach for these combinations - just incase they were typed in differently
	Part2 = "src=" + chr(34)
	Part3 = "src =" + chr(34)
	Part4 = "src = " + chr(34)
	
    d = LEN(Part2)
    
    b = 1
    do
	 a = instr(b, templateHTML, part1)
	 if a then 
		a = instr(a + 1, templateHTML, part2)
		if a = 0 then a = instr(a + 1, templateHTML, part3): d = LEN(Part3)
		if a = 0 then a = instr(a + 1, templateHTML, part4): d = LEN(Part4)
		IF a then 
			c = instr(a + d, templateHTML, chr(34))
			IF c - a - d - 1 > 1 then 
				filename = trim(mid(templateHTML, a + d, c - a - d))
				if lcase(filename) = "*|imagename|*" then 
					filename = SignatureName + "." + DefaultImageType
				end if
				OK = DownloadFile(filename, sourceUrl, destinationDirectory)
			END IF
			b = c + 1
		END IF
	  end if
    LOOP UNTIL b> LEN(templateHTML) OR a=0

END SUB


Function DownloadFile(ByVal filename, ByVal sourceUrl, ByVal destinationDirectory)
	
	' Dont download a file if running in Manual Mode... assumes the file is already in the %appdata% Signatures Folder
	IF ManualUser = True then 
		AddDebug "Using Existing File " & destinationDirectory & filename
		exit function
	END IF
	
	' Download a File
	AddDebug "Downloading FILE " & filename& " FROM " & sourceUrl & " TO " & destinationDirectory 
		
    Dim sourceFileUrl: sourceFileUrl = sourceUrl & filename
    Dim destinationFilePath: destinationFilePath = destinationDirectory & filename       
	
    Dim httpRequest: Set httpRequest = CreateObject("Msxml2.ServerXMLHTTP")

	
    'on error resume next
    httpRequest.open "GET", sourceFileUrl, false
    httpRequest.setRequestHeader "Cache-Control", "max-age=0"
    httpRequest.send()       
        
    if httpRequest.status = 200 Then
        Dim adoStream: Set adoStream = CreateObject("ADODB.Stream") 

        adoStream.Type = 1 'adTypeBinary
        adoStream.Open
		adoStream.Position = 0    'Set the stream position to the start
        adoStream.Write httpRequest.ResponseBody


        ' if the file already exists. remove it
        If fileSystem.FileExists(destinationFilePath) Then fileSystem.DeleteFile(destinationFilePath)       
		
        ' save a new copy of the file
        adoStream.SaveToFile(destinationFilePath)
        adoStream.Close

        Set adoStream = nothing
        DownloadFile = destinationFilePath
		AddDebug "Download File - OK"
    Else
        AddDebug "Download File - FAILED ("
        DownloadFile = ""

    End if

    on error goto 0
    Set httpRequest = Nothing

End Function


Function GetTPLFile(ByVal filename, ByVal sourceUrl)
	
	' Download a File
	AddDebug "Reading File : " & filename& " From : " & sourceUrl

    Dim sourceFileUrl: sourceFileUrl = sourceUrl & filename
	'Dim httpRequest : Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    'Dim httpRequest: Set httpRequest = CreateObject("Msxml2.ServerXMLHTTP")
	Dim httpRequest: Set httpRequest = CreateObject("Msxml2.ServerXMLHTTP.6.0")
	
	Dim txt
	
    'on error resume next
    httpRequest.open "GET", sourceFileUrl, false
    httpRequest.setRequestHeader "Cache-Control", "max-age=0"
    httpRequest.send()       
		
    if httpRequest.status = 200 Then
		AddDebug "Read File - OK"
		txt = httpRequest.responseText
        
    Else
        AddDebug "Read File - FAILED (" & cstr(httpRequest.status) &")"
        
    End if

    on error goto 0
    Set httpRequest = Nothing

End Function


Function IIf( expr, truepart, falsepart )
    IIf = falsepart
    If expr Then IIf = truepart
End Function


Function GetUserDN(Username, DomainName)

	' https://www.rlmueller.net/NameTranslateFAQ.htm
	
	on error resume next
	' Get the users Fully Qualified Domain Name
    Dim nameTranslate
    SET nameTranslate = CreateObject("NameTranslate")
	nameTranslate.Init ADS_NAME_INITTYPE_GC, ""
    nameTranslate.set ADS_NAME_TYPE_NT4, DomainName & "\" & Username
    GetUserDN = nameTranslate.Get(ADS_NAME_TYPE_1779)
	GetUserDN = Replace(GetUserDN, "/", "\/")
	on error goto 0
	
	Set nameTranslate = Nothing
	
End Function

Function AuthenticateDomain ( logonName, Password)

ldapFilter = "(samAccountName=" & logonName & ")"      'you could also search for an UPN here...

Set ado = CreateObject("ADODB.Connection")
ado.Provider = "ADSDSOObject"
ado.Properties("User ID") = logonName
ado.Properties("Password") = password
ado.Properties("Encrypt Password") = True
ado.Open "ADSearch" 
Set objectList = ado.Execute("<LDAP://" & ADserverIP  & "/" & ADDomain & ">;" & ldapFilter & ";distinguishedName,samAccountName,displayname,userPrincipalName;subtree")

While Not objectList.EOF
    AuthenticateDomain = objectList.Fields("distinguishedName")
    Username =  objectList.Fields("samAccountName")

    'On Error Resume Next 
    'displayName = "" : displayName = objectList.Fields("displayname")
    'logonNameUPN = "" : logonNameUPN = objectList.Fields("displayname")
    'On Error Goto 0

    'WScript.Echo logonName & " " & logonNameUPN  & " " & displayName & " " & userDN
   
    objectList.MoveNext
Wend


end function

Function isBetweenDate(BeginDate, EndDate)
	if IsEmpty(BeginDate) = False and IsEmpty(EndDate) = false then 
		If Date() >= BeginDate and Date() <= EndDate Then
			isBetweenDate = True
		Else
			isBetweenDate = False
		End If
	Else
		isBetweenDate = False
	end if

End Function

Function IsMember(ByVal objADObject, ByVal strGroup)

    ' Function to test for group membership.
    ' currentUserGroups is a dictionary object with global scope.
	if not objADObject is Nothing then 
		If (IsEmpty(currentUserGroups) = True) Then
			Set currentUserGroups = CreateObject("Scripting.Dictionary")
		End If
		If (currentUserGroups.Exists(objADObject.sAMAccountName & "\") = False) Then
			Call LoadGroups(objADObject, objADObject)
			currentUserGroups.Add objADObject.sAMAccountName & "\", True
		End If
		IsMember = currentUserGroups.Exists(objADObject.sAMAccountName & "\" & strGroup)
	END IF

End Function

Sub LoadGroups(ByVal objPriObject, ByVal objADSubObject)

    ' Recursive subroutine to populate dictionary object currentUserGroups.
	on error resume next
	
    Dim colstrGroups, objGroup, j

    currentUserGroups.CompareMode = vbTextCompare
    colstrGroups = objADSubObject.memberOf

    If (IsEmpty(colstrGroups) = True) Then
        Exit Sub
    End If

    If (TypeName(colstrGroups) = "String") Then
        ' Escape any forward slash characters, "/", with the backslash
        ' escape character. All other characters that should be escaped are.
        if instr(colstrGroups,"\/") = 0 then  colstrGroups = Replace(colstrGroups, "/", "\/")
				
        Set objGroup = GetObject("LDAP://" & colstrGroups)
        If (currentUserGroups.Exists(objPriObject.sAMAccountName & "\" & objGroup.sAMAccountName) = False) Then
            currentUserGroups.Add objPriObject.sAMAccountName & "\" & objGroup.sAMAccountName, True
            Call LoadGroups(objPriObject, objGroup)
        End If
        Set objGroup = Nothing
        Exit Sub
    End If

    For j = 0 To UBound(colstrGroups)
        ' Escape any forward slash characters, "/", with the backslash
        ' escape character. All other characters that should be escaped are.
        if instr(colstrGroups(j),"\/") = 0 then colstrGroups(j) = Replace(colstrGroups(j), "/", "\/")
        Set objGroup = GetObject("LDAP://" & colstrGroups(j))
        If (currentUserGroups.Exists(objPriObject.sAMAccountName & "\" & objGroup.sAMAccountName) = False) Then
            currentUserGroups.Add objPriObject.sAMAccountName & "\" & objGroup.sAMAccountName, True
            Call LoadGroups(objPriObject, objGroup)
        End If
    Next
    Set objGroup = Nothing

End Sub

Function CreateIDNumber(strUserName)
	
	' Create a user unique number based on the users username
	DIM j, ch, ID
	ID=""

	if len(strUserName) > 0 then 
		for j = 1 to len(strUserName)
			ch=hex(asc(mid(strUserName,j,1)))
			ID=ID+cstr(ch)
		Next
		'AddDebug "IDnumber : " & ID
	end if
	CreateIDNumber = ID
	
End Function

Sub LoadLastDetails(ByVal signaturesFolderPath)

	' Load users details from a previously saved file
	Dim  LoadDetails, LoadFileName, strInfo
	LoadFileName = signaturesFolderPath & "info.bak"
	AddDebug "Reading Previous Details Information from : " & SaveFileName
	on error resume next
	Set LoadDetails = fileSystem.OpenTextFile(LoadFileName, 1, False)
	strInfo=LoadDetails.ReadAll
	LoadDetails.Close
	on error goto 0
	
	Set LoadDetails = Nothing
	
End Sub

Sub SaveLastDetails(ByVal signaturesFolderPath)

	' Save users details so they can be checked next time and know if somthing has changed.
	Dim  SaveDetails, SaveFileName, strInfo
	SaveFileName = signaturesFolderPath & "info.bak"
	AddDebug "Saving Current Detais Information to : " & SaveFileName
	on error resume next
	Set SaveDetails = fileSystem.CreateTextFile(SaveFileName, True)
	SaveDetails.Write(strInfo)	
	SaveDetails.Close
	on error goto 0
		
	Set SaveDetails = Nothing
	
End Sub

Sub CreateVcard(ByVal signaturesFolderPath, ByVal SignatureName)

	' Create Microsoft Vcard details based on the users details.
	' This vcard info is not used currently. trying to add it to outlook is not working.
	
	Dim VCardFile,vCard

	vCard ="BEGIN:VCARD" + VbCrLf +"VERSION:2.1" + VbCrLf
	if len(firstname) then vCard=vCard +"N:"+StripHTML(firstname)+";"+StripHTML(lastname) + VbCrLf
	if len(name) then vCard=vCard +"FN:"+StripHTML(name) + VbCrLf
	if len(company) then vCard=vCard +"ORG:" + StripHTML(company) + VbCrLf
	if len(title) then vCard=vCard +"TITLE:" + StripHTML(title) + VbCrLf
	if len(phone) then vCard=vCard +"TEL;WORK;VOICE:" + StripHTML(phone) + VbCrLf
	if len(mobile) then vCard=vCard +"TEL;CELL;VOICE:" + StripHTML(mobile) + VbCrLf
	if len(companyfax) then vCard=vCard +"TEL;WORK;FAX:" + StripHTML(companyfax) + VbCrLf
	if len(webpage) then vCard=vCard +"URL;WORK:" +StripHTML(webpage) + VbCrLf
	if len(address) then 
		vCard=vCard + "ADR;WORK;PREF:"+StripHTML(address)+";;"+StripHTML(pobox)+";"+StripHTML(suburb)+";"+StripHTML(state)+";"+StripHTML(postcode)+";"+StripHTML(countryname)+ vbCrlf 
		vCard=vCard +"LABEL;WORK;PREF;ENCODING=QUOTED-PRINTABLE:" + StripHTML(address)+"=0D=0A=" + VbCrLf
		if len(suburb) then vCard=vCard + StripHTML(suburb)
		if len(state) then vCard=vCard + ", " + StripHTML(state)
		if len(countryname) then vCard=vCard + ", " +StripHTML(countryname)
		if len(postcode) then vCard=vCard + ", " + StripHTML(postcode)
		vCard=vCard + VbcrLf
		vCard = vCard + "X-MS-OL-DEFAULT-POSTAL-ADDRESS:2" + VbCrLf
	end if
	if len(email) then vCard=vCard +"EMAIL;PREF;INTERNET:" + StripHTML(email) + VbCrLf
	if len(vcardPhoto) then vCard=vCard +"PHOTO;TYPE=JPEG;ENCODING=BASE64:" + VbCrLf + vcardPhoto + VbCrLf+ VbCrLf
	vCard=vCard + "X-MS-OL-DESIGN;CHARSET=utf-8:<card xmlns=""http://schemas.microsoft.com/office/outlook/12/electronicbusinesscards"" ver=""1.0"" layout=""left"" bgcolor=""ffffff""><img xmlns="""" align=""fit"" area=""16"" use=""cardpicture""/>" +VbCrLf
	VCard=VCard + "<fld xmlns="""" prop=""name"" align=""left"" dir=""ltr"" style=""b"" color=""000000"" size=""10""/><fld xmlns="""" prop=""org"" align=""left"" dir=""ltr"" color=""000000"" size=""8""/><fld xmlns="""" prop=""title"" align=""left"" dir=""ltr"" color=""000000"" size=""8""/><fld xmlns="""" prop=""telwork"" align=""left"" dir=""ltr"" color=""d48d2a"" size=""8""><label align=""right"" color=""626262"">Work</label></fld><fld xmlns="""" prop=""telcell"" align=""left"" dir=""ltr"" color=""d48d2a"" size=""8""><label align=""right"" color=""626262"">Mobile</label></fld><fld xmlns="""" prop=""email"" align=""left"" dir=""ltr"" color=""d48d2a"" size=""8""/><fld xmlns="""" prop=""addrwork"" align=""left"" dir=""ltr"" color=""000000"" size=""8""/><fld xmlns="""" prop=""webwork"" align=""left"" dir=""ltr"" color=""000000"" size=""8""/><fld xmlns="""" prop=""blank"" size=""8""/><fld xmlns="""" prop=""blank"" size=""8""/><fld xmlns="""" prop=""blank"" size=""8""/><fld xmlns="""" prop=""blank"" size=""8""/><fld xmlns="""" prop=""blank"" size=""8""/><fld xmlns="""" prop=""blank"" size=""8""/><fld xmlns="""" prop=""blank"" size=""8""/><fld xmlns="""" prop=""blank"" size=""8""/></card>" + VbCrLf
	vCard=vCard +"REV:VCARD" + VbCrLf
	vCard=vCard +"END:VCARD"

	if len(SignatureName) then 
		AddDebug "Creating Outlook contact (vCard) file : " & signaturesFolderPath & SignatureName & ".vcf"
		on error resume next
		Set VCardFile = fileSystem.CreateTextFile(signaturesFolderPath & SignatureName & ".vcf", True)
		VCardFile.Write(vCard)	
		VCardFile.Close
	end if	
	
	Set VCardFile = Nothing

END Sub

Function TransformText(TransformType, byval str)
	
	TransformText=eval(str)
	
	if len(TransformType) > 0 then 
		select case  ucase(TransformType)
		case "COMPANYPHONE"
			TransformText=PhoneNumber(eval(str))	
		case "INTCOMPANYPHONE"
			TransformText=PhoneINTNumber(eval(str))	
		case "PHONE"
			TransformText=PhoneNumber(eval(str))	
		case "MOBILE"
			TransformText=MobileNumber(eval(str))	
		case "INTPHONE"
			TransformText=PhoneINTNumber(eval(str))	
		case "INTMOBILE"
			TransformText=MobileINTNumber(eval(str))	
		case "UPPER"
			TransformText=ucase(eval(str))
		case "LOWER"
			TransformText=lcase(eval(str))
		case "PROPER"
			TransformText=ProperCase(eval(str))
		case "FULLSTATE"
			TransformText=FullState(eval(str))
		case "SHORTSTATE"
			TransformText=ShortState(eval(str))
		case "FULLCOUNTRY"
			TransformText=FullCountry(eval(str))
		case "SHORTCOUNTRY"
			TransformText=ShortCountry(eval(str))
		End Select
	AddDebug "Mod Variable: " & str &" '" & TransformType & "' -> " & TransformText
	end if
		
END Function


Function ProperCase(byval Inn)

	Dim xx, Tr, Txt, Strt, Tr2, Bad, Tr3
	Txt = Inn
	'Txt = StrConv(Txt, vbProperCase)
	if len(Txt) > 0 then 
		Txt = ltrim(Txt)
		Txt = ucase(left(Txt,1)) + mid(Txt,2,len(Txt))
	end if
	
	'---------------O'Neal Etc------------
	Strt = 1
	Do
		xx = InStr(Strt, Txt, "'")
		If xx = 0 Then Exit Do
		Strt = xx + 1
		If xx < Len(Txt) And xx > 1 Then
			If xx = 2 Then
				Tr3 = Mid(Txt, xx + 1, 1)
				If Tr3 <> " " Then
					Tr = Mid(Txt, xx + 1, 1)
					Mid(Txt, xx + 1, 1) = UCase(Tr)
				End If
			Else
				Tr2 = Mid(Txt, xx - 2, 1)
				Tr3 = Mid(Txt, xx + 1, 1)
				If Tr2 = " " And Tr3 <> " " Then
					Tr = Mid(Txt, xx + 1, 1)
					Mid(Txt, xx + 1, 1) = UCase(Tr)
				End If
			End If
		End If
	Loop
'--------Hyphens---------
Strt = 1
Do
   xx = InStr(Strt, Txt, "-")
   If xx = 0 Then Exit Do
   Strt = xx + 1
   If xx < Len(Txt) Then
      Tr = Mid(Txt, xx + 1, 1)
      Mid(Txt, xx + 1, 1) = UCase(Tr)
   End If
Loop
'-------------------------Mc-----------
Strt = 1
Do
   xx = InStr(Strt, Txt, "Mc")
   If xx = 0 Then Exit Do
   Strt = xx + 2
   If xx + 1 < Len(Txt) Then
      Tr = Mid(Txt, xx + 2, 1)
      Mid(Txt, xx + 2, 1) = UCase(Tr)
   End If
Loop
	'-------------------------Mac-----------
	Strt = 1
	Bad = ":he:hi:hr:hz:ho:hs:ro:ru:ra:ac:ad:aq:aw:ar:ab:ki:kl:ks:"
	Bad = Bad + "in:ul:um:er:es:ed:ke:ca:co:le:on:"
	Do
		xx = InStr(Strt, Txt, "Mac")
		If xx = 0 Then Exit Do
		Strt = xx + 3
		If xx + 6 < Len(Txt) Then
			Tr2 = Mid(Txt, xx + 3, 2)
			If not(Instr(Tr2, " ")) then
				Tr2 = ":" + Tr2 + ":"
				If not (Instr(Bad, Tr2)) then
					Tr = Mid(Txt, xx + 3, 1)
					Mid(Txt, xx + 3, 1) = UCase(Tr)
				end if
			end if
		End If

	Loop
	'-----------------
	ProperCase = Txt
	
End Function


Function PhoneNumber (byval strPhone)
	
	DIM a, tmpNumber, newnum,fsp
	tmpNumber=strPhone
	' Check for International Number
	if len(tmpNumber) > 4 then 
		if left(tmpNumber,3)="+61" then 
			tmpNumber=mid(tmpNumber,4,len(tmpNumber))
			tmpNumber="0" + tmpNumber
		end if
	end if
	
	tmpNumber=replace(tmpNumber," ","")
	tmpNumber=replace(tmpNumber,"-","")	
	
	if len(tmpNumber) = 10 then 
		newnum = tmpNumber
	    if left(tmpNumber,4) ="1300" then newnum = left(tmpNumber,4)+" "+mid(tmpNumber,5,3)+" "+mid(tmpNumber,8,3)
		if left(tmpNumber,4) ="1800" then newnum = left(tmpNumber,4)+" "+mid(tmpNumber,5,3)+" "+mid(tmpNumber,8,3)
		if left(tmpNumber,1)="0" then newnum = left(tmpNumber,2)+" " + mid(tmpNumber,3,4)+" " +mid(tmpNumber, 7,4)
		PhoneNumber = newnum
	else
	    if left(tmpNumber,2) ="13" then 
			newnum = left(tmpNumber,2)+" "+mid(tmpNumber,3,2)+" "+mid(tmpNumber,5,2)
			PhoneNumber = newnum	
		else	
			PhoneNumber = strPhone
		end if
	end if
	
end function

Function PhoneINTNumber (byval strPhone)
	
	DIM a, tmpNumber, newnum
	tmpNumber=replace(strPhone," ","")
	tmpNumber=replace(tmpNumber,"-","")
	if len(tmpNumber) = 10 then 
		newnum = tmpNumber
	    if left(tmpNumber,1)="0" then newnum = InternationalPrefix+" " +mid(tmpNumber,2,1)+" " + mid(tmpNumber,3,4)+" " +mid(tmpNumber, 7,4)
		PhoneINTNumber = newnum
	else
		PhoneINTNumber = strPhone
	end if
	
end function


Function MobileNumber (byval strMobile)
	DIM a, tmpNumber, newnum
	tmpNumber=replace(strMobile," ","")
	tmpNumber=replace(tmpNumber,"-","")
	
	if len(tmpNumber) = 10 then 
		newnum = tmpNumber
		if left(tmpNumber,1)="0" then newnum = left(tmpNumber,4)+" " + mid(tmpNumber,5,3)+" " +mid(tmpNumber, 8,3)
		MobileNumber = newnum
	else
		MobileNumber = strMobile
	end if
	
end function

Function MobileINTNumber (byval strMobile)
	DIM a, tmpNumber, newnum
	tmpNumber=replace(strMobile," ","")
	tmpNumber=replace(tmpNumber,"-","")
	if len(tmpNumber) = 10 then 
		newnum = tmpNumber
		if left(tmpNumber,1)="0" then newnum = InternationalPrefix+" "+mid(tmpNumber,2,3)+" " + mid(tmpNumber,5,3)+" " +mid(tmpNumber, 8,3)
		MobileINTNumber = newnum
	else
		MobileINTNumber = strMobile
	end if

	
end function

Function FullCountry (byval strCountry)
	
	DIM a
	a=0
	DO
	  IF lcase(CountryCodes(a)) = lcase(strCountry) Then
	     FullCountry = CountryNames(a)
	     EXIT Do
	  END IF
	a=a+1
	loop until a> ubound(CountryCodes)

END Function

Function ShortCountry (byval strCountry)
	
	DIM a
	a=0
	DO
	  IF lcase(CountryNames(a)) = lcase(strCountry) Then
	     ShortCountry = CountryCodes(a)
	     EXIT Do
	  END IF
	a=a+1
	loop until a> ubound(CountryCodes)

END Function


' Returns Full State Name from Short State Name
Function FullState (byval strState)
	
	DIM fState
	fState = strState

	SELECT Case ucase(strState)
	case "QLD","QUEENSLAND"
		fState ="Queensland"
	case "VIC","VICTORIA"
		fState ="Victoria"
	case "NSW","NEW SOUTH WALES","NEWSOUTH WALES","NEWSOUTHWALES"
		fState ="New South Wales"
	case "SA","SOUTH AUS","SOUTH AUSTRALIA"
		fState ="South Australia"
	case "TAS","TASMANIA"
		fState ="Tasmania"
	case "ACT","AUSTRALIAN CAPITAL","AUSTRALIAN CAPITAL TERRITORY"
		fState ="Australian Capital Territory"
	case "JBT","JERVIS BAY"," JERVIS BAY TERRITORY"
		fState ="Jervis Bay Territory"
	case "NT","NORTHAN TERRITORY"
		fState ="Northan Territory"
	END SELECT
	
	FullState = fState
	
END Function


' Returns Short State Name from Full State Name
Function ShortState (byval strState)
	
	DIM fState
	fState = strState

	SELECT Case lcase(strState)
	case "qeensland","qld"
		fState = "QLD"
	case "victoria","vic"
		fState ="VIC"
	case "new south wales","nsw"
		fState ="NSW"
	case "south australia","sa"
		fState ="SA"
	case "tasmania","tas"
		fState ="TAS"
	case "australian capital territory","act"
		fState ="ACT"
	case "jervis bay territory","jbt"
		fState ="JBT"
	case "northan territory","nt"
		fState ="NT"
	END SELECT
	ShortState = fState
	
END Function


Function CheckExtra(byval strGroupDN)


	DIM wshShell, objSysInfo, user, domainName, userLDAP, userDN, strNetBIOSDomain
	Set objSysInfo = CreateObject( "WinNTSystemInfo" )
	Set wshShell = CreateObject( "WScript.Shell" )
	userName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
	strNetBIOSDomain = wshShell.Environment("Process").Item("userdomain")
	
    SET user = CreateObject("WScript.Network")
    userName = user.UserName
    domainName = user.UserDomain
	'domainName = objSysInfo.DomainName
	'WScript.Echo UserName, domainName
	
	UserDN = GetUserDN(userName,domainName)
	
    if len(UserDN)> 0  then 
		DIM objGroup, objMember
		SET objGroup = GetObject(strGroupDN)
		For Each objMember in objGroup.Members
			If (objMember.distinguishedName = UserDN) Then
				extranote=objGroup.Description
				Set objGroup = Nothing
				Exit Function
			End If
		Next
		SET objGroup = Nothing
	end if	
	
	SET user = Nothing
	SET objSysInfo = Nothing
	
	AddDebug "LDAP: User " & UserDN &" not found in MemberNote"
	
End Function

Function SetDefaultMailSignature(signatureName)

	if AdminDefaultSignature = False Then

		on error resume next
		' Init local vars
		Dim winWordOpened: winWordOpened = False
		Dim winWord: Set winWord = Nothing

		' is word already open?
		If winWord Is Nothing Then
			Set winWord = CreateObject("Word.Application")
			winWordOpened = True
		Else
			Set winWord = GetObject(, "Word.Application")
		End If
   
		' Set the signatures to default
		Dim outlookEmailOptions: Set outlookEmailOptions = winWord.EmailOptions
		Dim signatureObjects: Set signatureObjects = outlookEmailOptions.EmailSignature
		signatureObjects.NewMessageSignature = signatureName
		signatureObjects.ReplyMessageSignature = signatureName
		AddDebug ""
		AddDebug "Setting Signature " + signatureName +" as Default."
		AddDebug ""
		' if we opened word, then close it...
		If winWordOpened Then winWord.Quit
		Set winWord = Nothing
		on error goto 0
		
	else
		AddDebug "Overriding SetDefaultSignature using '-default' in users notes field."
	end if
	
End Function

Function StripHTML(byval strText)
	on error resume next
	Dim tmpStr
	Dim RegEx
	Set RegEx = New RegExp

	RegEx.Pattern = "<[^>]*>"
	RegEx.Global = True

	tmpStr = RegEx.Replace(strText, "")
	tmpStr = Replace(tmpStr, Spc, "")
	
	StripHTML = tmpStr
	on error goto 0
END Function

Function CheckVariables(templateHTML)
	dim FindVariable,a, FindBracket, FindEnd
	IF instr(templateHTML,"<*|Var(") then 
		a=1
		do
		FindVariable = instr(a, templateHTML,"<*|Var(")
		if FindVariable then 
			FindBracket=instr(FindVariable+6,templateHTML,")|*")
			if FindBracket then 
				FindEnd=instr(FindBracket+3,templateHTML,">")
				if FindEnd and MaxVariables < 101 then 
					MaxVariables = MaxVariables + 1
					Variable(MaxVariables) = mid(templateHTML, FindVariable +8, FindBracket - FindVariable -7)
					VariableValue(MaxVariables) = mid(templateHTML, FindBracket +4, FindEnd - FindBracket -3)
					AddDebug "HTML Variable: " & Variable(MaxVariables) &" value : " & VariableValue(MaxVariables)
				a=FindEnd +1
				end if
			end if
		end if
		loop until FindVariable = 0 
	END IF
	
END Function

Function SetVariables(templateHTML)
	dim FindVariable,a, FindBracket, FindEnd
	
	IF MaxVariables > 0 then 
		a=1
		do
			FindVariable = instr(templateHTML,Variable(a))
			if FindVariable then 
				FindBracket=instr(FindVariable+len(Variable(a))+1,templateHTML,":")
				if FindBracket then 
					FindEnd=instr(FindBracket+1,templateHTML,";")
					if FindEnd then 
						templateHTML= left(templateHTML, FindVariable+len(Variable(a))) + VariableValue(MaxVariables) + mid(templateHTML, FindEnd,len(templateHTML))
						AddDebug "Set Variable: " & Variable(MaxVariables) &" 'HTML' -> " & VariableValue(MaxVariables)
					end if
				end if
			end if
			a=a+1
		loop until a > MaxVariables
	END IF
	
END Function

Function CheckVar(byval strText)
	dim FindVar,a , FindClose, FindVal
	a=1
	if len(strText) > 0 then 
		CheckVar= strText
		' Find all the character values in the text 
		if instr(a, lcase(strText),"char(") then 
			CheckVar= ""
			do
				FindVar=instr(a, lcase(strText),"char(")
				if FindVar then 
					FindClose = instr(FindVar+7,lcase(strText),")")
					if FindClose then 
						FindVal= cint(mid(strText,FindVar+5,FindClose-FindVar-5))
						if FindVal > 255 then FindVal =0
						Checkvar=Checkvar+chr(FindVal)
					end if
					a=FindVar+1
					'AddDebug "Found char("+cstr(FindVal)+") value in HTML variable"
				end if
			loop until FindVar =0 or a => len(strText)
		end if
	end if
	
END Function


Function Encode64(byval strFileName)

	Dim Cobj, CElem
	Dim inStream, fileBytes
	Set inStream = CreateObject("ADODB.Stream")
	inStream.Open
	inStream.type = 1
	inStream.LoadFromFile strFileName
	fileBytes=InStream.Read()

	Set CObj = CreateObject("Microsoft.XMLDOM")
	Set CElem = Cobj.createElement("tmp")
	CElem.DataType = "bin.base64"
	CElem.NodeTypedValue = fileBytes
	Encode64 = CElem.Text
	
	AddDebug "Encoded Base64: " & strFileName

	set inStream = Nothing
	set Cobj = Nothing
	set CElem = Nothing

End Function


Function FindStartDayinMonth(byval dteDate) 
	Dim FirstDayOfMonth
	Dim DayModifier
	DayModifier=cint(left(dteDate,2))
    FirstDayOfMonth = DateSerial(Year(dteDate), Month(dteDate), 1)
	if Weekday(FirstDayOfMonth) = 1 then 
		DayModifier = DayModifier + 1
	end if
	if Weekday(FirstDayOfMonth) = 7 then 
		DayModifier = DayModifier + 3
	end if
	FindStartDayinMonth = ConvertToDate(cstr(DayModifier)+mid(dtedate,3,len(dtedate)))
END FUNCTION 

Function FindWeekinDate(byval dteDate, byval weeknum ) 
	
	DIM FMonday
	
    Dim MondayModifier 
    Dim FirstDayOfMonth
        
    FirstDayOfMonth = DateSerial(Year(dteDate), Month(dteDate), 1)

    MondayModifier = (9 - Weekday(FirstDayOfMonth)) Mod 7
    FMonday = (DateAdd("d", MondayModifier+((weeknum *7)-7), FirstDayOfMonth))
	FMonday = ConvertToDate(FMonday)
	
	' Check if Monday is New Years Day, if it is then the start day is Tuesday
	IF left(FMonday, 5) = "03/01"  Then
		FMonday = "04/01/" + right(FMonday, 4)
	END IF
	IF left(FMonday, 5) = "02/01"  Then
		FMonday = "03/01/" + right(FMonday, 4)
	END IF
	IF left(FMonday, 5) = "01/01"  Then
		FMonday = "02/01/" + right(FMonday, 4)
	END IF
	
	FindWeekinDate = FMonday
	
End Function

Function FindSundayinDate(byval dteDate, byval weeknum ) 
	
	DIM FSunday
	
    Dim SundayModifier 
    Dim FirstDayOfMonth
        
    FirstDayOfMonth = DateSerial(Year(dteDate), Month(dteDate), 1)

    SundayModifier = (8 - Weekday(FirstDayOfMonth)) Mod 7
    FSunday = (DateAdd("d", SundayModifier+((weeknum *7)-7), FirstDayOfMonth))
	FindSundayinDate = ConvertToDate(FSunday)
	
End Function

Function DateAddDay(byval dteDate, byval intdays)

	DIM Modifyer
	Modifyer = DateAdd("d", intdays, dteDate)
	DateAddDay = ConvertToDate(Modifyer)

End Function

Function FindLastDay(byval dteDate, byval extradays)

DIM ChrDay, ChrWeekDay, Modifyer

ChrDay = cstr(25 - extradays)+"/12/" + right(dteDate, 4)

ChrWeekDay = Weekday(ChrDay)
Modifyer = dteDate

' If the Christmas day is on a Weekend then adjust the last day backwards so the friday is a holiday
select case ChrWeekDay
	case "1" ' Sunday
	Modifyer = DateAdd("d", -3, ChrDay) ' Thursday
	case "2" ' Monday
	Modifyer = DateAdd("d", -4, ChrDay) ' Friday
	case "3" ' Tuesday
	Modifyer = DateAdd("d", -5, ChrDay) ' Friday
	case "4" ' Wednesday
	Modifyer = DateAdd("d", -6, ChrDay) ' Friday 
	case "5" ' Thursday
	Modifyer = DateAdd("d", -2, ChrDay) ' Tuesday
	case "6" ' Friday
	Modifyer = DateAdd("d", -2, ChrDay) ' Wednesday
	case "7" ' Saturday
	Modifyer = DateAdd("d", -2, ChrDay) ' Thursday

END Select

FindLastDay = ConvertToDate(Modifyer)

END Function

Function th(intDay)
	DIM str 
	str = cstr(IntDay)
	if len(str) > 0 then 
		select case intDay
		case 1, 21, 31
			str = str + "st"
		case 2, 22
			str = str + "nd"
		case 3,23
			str = str + "rd"
		case Else
			str = str + "th"
		end select
	END IF
	th = str
END Function

Function ConvertToDate(strDate)
	
	DIM find, a, d, m, y, ad, am, Cd
	
	ConvertToDate = strDate
	
	' Find a "/" or "-" in the date string
	find = "/"
	a = instr(strDate, find)
	IF a = 0 then 
		find = "-"
		a = instr(strDate, find)
	END IF
	IF a >  0 then 
		ad = a
		d = left(strDate, ad - 1)
		am = instr(ad + 1, strDate, Find)
		IF am > 0 then 
			m = mid(strDate, ad + 1, (am - ad)-1)
			y = mid(strDate, am + 1,len(strDate))
		END IF
		if len(d) > 0 and len(m) > 0 and len(y) = 4 then 
			d = PadString(d, 2, "0", "left")
			m = PadString(m, 2, "0", "left")
			ConvertToDate = d + "/" + m + "/" + y
		END IF
	END IF
	
End Function

Function PadString(pString,pLength,pChar,pSide)
 
  DIM strString
   
  strString = pString
   
  strPadding = String(pLength,pChar)  
   
  If lcase(pSide)="left" then
    strString = strPadding & strString
    strString = Right(strString,pLength)
  else
    strString = strString & strPadding
    strString = Left(strString,pLength)  
  End if
   
  PadString = strString 
    
End Function

Function DateExpanded(strDate)
	strDate = ConvertToDate(strDate)
	if len(strDate) = 10 then 
		DateExpanded = WeekDayName(Weekday(strDate)) + " " & th(Day(strDate)) + " of " + MonthName(Month(strDate)) + " " + cstr(Year(strDate))
	Else
		DateExpanded = ""
	end if
End Function

Function DateExpanded2(strDate)
	strDate = ConvertToDate(strDate)
	if len(strDate) = 10 then 
		DateExpanded2 = MonthName(Month(strDate)) + " "+ cstr(Day(strDate)) + ", " + cstr(Year(strDate))
	Else
		DateExpanded2 = ""
	end if
End Function


Function REGKeyValue(strKeyPath)
	on error resume next
	REGKeyValue =  ""
	dim iValue
	iValue = cstr(shell.RegRead(strKeyPath))
	if len(iValue) > 0 then REGKeyValue = iValue
	on error goto 0	
End Function

function Ping(strComputer)
	
    dim objSh,objPing
    dim strPingOut, flag

	' Remove any NON DNS Chars leaving domain.name.anything to Ping
	strComputer = Replace(strComputer,"HTTPS","")
	strComputer = Replace(strComputer,"HTTP","")
	strComputer = Replace(strComputer,"FTP","")
	strComputer = Replace(strComputer,"://","")
	strComputer = Replace(strComputer,"/","")
	
	on error resume next
    set objPing = shell.Exec("%comspec% /c ping -4 -n 1 " & strComputer)
	strPingOut = objPing.StdOut.ReadAll
	
    if instr(LCase(strPingOut), "reply") then
        flag = TRUE
    else
        flag = FALSE
    end if

    on error goto 0	
	Ping = flag
	
	set objSh = Nothing
	set objPing = Nothing
	
end function


function DNSLookup(strComputer)

	dim objSh,nslookupObj, strName
    dim strNslookup, result, found, strNSLine, strNsErr
	dim f1, f2, F3
	
	found = 0
	result = ""
	strName = strComputer
	
	' Remove any NON DNS Chars leaving domain.name.anything to Ping
	strName = Replace(strName,"HTTPS","")
	strName = Replace(strName,"https","")
	strName = Replace(strName,"HTTP","")
	strName = Replace(strName,"http","")
	strName = Replace(strName,"FTP","")
	strName = Replace(strName,"ftp","")
	strName = Replace(strName,"://","")
	strName = Replace(strName,"/","")
	
	'on error resume next
	set nslookupObj = shell.Exec("%comspec% /c nslookup " & strName)
	strNsErr = nslookupObj.Stderr.ReadAll
	strNslookup = nslookupObj.StdOut.ReadAll

	if instr(strNsErr,"***") = 0 then 
		f1 = instr(1, strNslookup,"Address:  ")
		IF f1 > 0 then 
			f2 = instr(f1+1, strNslookup,"Address:  ")
			IF f2 > 0 then 
				F3 = instr(F2+1, strNslookup, VbCrLf)
				if F3 > 0 then 
					result = trim(mid(strNslookup, F2+10,F3-F2-10))
				end if
			END IF
		END IF
	END IF
	if result ="" then 
		AddDebug("NSLOOKUP " & strName & " | " & strNsErr & " | " & strNslookup & " | " & result)
	Else
		AddDebug("NSLOOKUP " & strName & " " & result & " - OK")
	END IF
	
    on error goto 0	
	DNSLookup = result
	
	set objS = Nothing
	set objPing = Nothing
	
end function


FUNCTION ReadUserInfoFile (strFileName)

	AddDebug "Reading " & strFileName
		
	DIM GName, UserInfoFile, UserInfo
	SET UserInfoFile = filesystem.OpenTextFile(strFileName, 1, False)
	USerInfo = UserInfoFile.ReadAll
	UserInfoFile.Close
	
	GName = ReadINISetting("GroupName", USerInfo,1)
	SignatureGroupFileName(1) = GName
	SignatureGroupName(1) = GName
	GName = ReadINISetting("GroupName", USerInfo,2)
	if len(GName) > 0 then 
		SignatureGroupFileName(2) = GName
		SignatureGroupName(2) = GName
	END IF
	UserName = ReadINISetting("UserName", USerInfo,1)
	DomainName = ReadINISetting("Domain", USerInfo,1)
	name = ReadINISetting("FullName", USerInfo,1)
	firstname = ReadINISetting("FirstName", USerInfo,1)
	lastname = ReadINISetting("LastName", USerInfo,1)
	title = ReadINISetting("Title", USerInfo,1)
	mobile = ReadINISetting("Mobile", USerInfo,1)
	phone = ReadINISetting("Phone", USerInfo,1)
	address = ReadINISetting("Address", USerInfo,1)
	email = ReadINISetting("Email", USerInfo,1)
	pobox = ReadINISetting("POBox", USerInfo,1)
	state = ReadINISetting("State", USerInfo,1)
	city = ReadINISetting("City", USerInfo,1)
	suburb = ReadINISetting("Suburb", USerInfo,1)
	country = ReadINISetting("Country", USerInfo,1)
	postcode = ReadINISetting("PostCode", USerInfo,1)
	office = ReadINISetting("Office", USerInfo,1)
	webpage = ReadINISetting("Webpage", USerInfo,1)
	countryname = ReadINISetting("CountryName", USerInfo,1)
	ipphone = ReadINISetting("IPphone", USerInfo,1)
	notes = ReadINISetting("Notes", USerInfo,1)

	ADAddress = address
	ADTitle = title
	ADCompany = company
	ADPhone = phone
	ADMobile = mobile
	ADstate = state
	ADpostcode = postcode
		 
	Set UserInfoFile = Nothing
	Set fSystem = Nothing

END Function

FUNCTION ReadINISetting(findstr, strText, occurance)

	DIM strfound, a, b,c,d, o, lfindstr, lstrText, nexto
	if occurance = 0 then occurance = 1
	O = 1
	nexto = 1
	if len(findstr) > 0 and len(strText) > 0 then 
		lfindstr = lcase(findstr)
		lstrText =lcase(strText)
		d=len(findstr)
		do
			a = instr(nexto, lstrText, lfindstr & " = "): c = 3
			if a = 0 then a = instr(nexto, lstrText, lfindstr & " ="): c = 2
			if a = 0 then a = instr(nexto, lstrText, lfindstr & "= "): c = 2
			if a = 0 then a = instr(nexto, lstrText, lfindstr & "=" ): c = 1
			if a > 0 and a < len(strText) and o = occurance then 
				b = instr((nexto-1) + a + c+ d, strText, chr(13))
				if b = 0 then b = len(strText)
				strFound = trim(mid(strText, a+c+d, b-a-d-c))
				AddDebug "INI Setting " & findstr & " -> " & strFound
			Else
				nexto = nexto+a+c+d+1
			END IF
		o = o + 1
		loop until o > occurance
	END IF
	
	ReadINISetting = strFound
	
END Function
