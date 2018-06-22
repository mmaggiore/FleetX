<%

'	This script assumes that the Session contains the following Crystal Report Engine
'	Objects:
'
'	"oApp" - Crystal Report Engine Application Object
'	"oRpt" - Crystal Report Engine Report Object
'	"oPageEngine - Crystal Report Engine Page Engine Object
'	HTML_FRAME viewer only
'	"tabArray" -  Array used to keep drilldown tab information
'	"CurrentPageNumber" - Current Page Number requested.
'	"lastknownpage" - Previous page number requested.
'	"LastPageNumber" - Last known page number requested.
'	Note:  Before creating the PageEngine object, call ReadRecords on the 
'	report object to that all the database records have been read.
'
'	Modifications
' 05/02/98
' Added the following features:
' Tab Query String Parameter
'	- This is the selected HTML_FRAME viewer tab's tabArray index value.  
' Page Expiry Time
'	-  The page will expire when downloaded by browser so that user is insured that all data
' will be current.
' DrillDown Tabs
'	- Added in the session("tabArray") object to keep track of the HTML_FRAME drill down tabs.
' RFSH Query String Parameter
'	- Database is refreshed and ErrorValue 0 is sent to Java and active X viewers on success.
'This causes viewers to NOT refresh browser window.
' SRCH Query String Parameter and HTML_FRAME Viewer
'	-  Added javascript window.alert function call to indicate when text is not found in rpt view.

' 09/08/98
' Added the following features:
'	- cmd handling for map_dd(Map Drill Down) and export(Exporting reports from viewers).
'	- PageGenerator object creation for Out of Place Subreports.  Modification was made to RetrieveObjects procedure.
'	- Modified the get_ttl handling to call the PageGenerator object's RenderTotallerETF method 
'		instead of the PageEngine's method.  This was done to support the group by feature.
' 08/03/99
' Added the following features
'	- cmd/rfsh handling for pages with and without place holders.
'	- cmd/rfsh handling for pages requiring and not requiring total page count.


 On Error Resume Next

'  The oEMF object is a helper object to create EMFs (Ecapsulated Messages) for the viewers.
'  The viewers use EMFs to display errors and navigate to specific pages of the report.

If Not IsObject(session("oEMF")) then
	Set session("oEMF") = Server.CreateObject("CREmfgen.CREmfgen.1")
	Call CheckForError
End if


'	Initialize all Global variables
'	These will contain the page generator and page collection

	Dim goPageGenerator		' page generator object
	Dim goPageCollection	' page collection object
	Dim goPageGeneratorDrill' page generator object in Drill Down Context
	Dim goPage				' the page object
	Dim gvGroupPathDD		' drill down group path, this is an array.
	Dim gvGroupPath			' this is branch, aka Group Path converted from string passed on the QS, it is an Array
	Dim gvGroupLevel		' this is the Group level, converted from the string passed on the QS, it is an Array
	Dim gvMaxNode			' this represents the number of nodes to retrieve for the totaller, it is set to an empty array
	Dim gvTotallerInfo		' this represents the group path of the requested totaller.
	Dim glX					' this is the X Coordinate for a drill down on a graph or Map
	Dim glY					' this is the Y Coordinate for a drill down on a graph or Map
	Dim gvPageNumber		' holds the requested page number
	Dim gvURL				' URL to redirect to
	Dim gsErrorText			' holds the error text to be sent to the viewer.
	Dim ExportOptions		' Export Options Object 
	Dim slX					' this is the X Coordinate for a selection of Out of Place subreport
	Dim slY					' this is the Y Coordinate for a selection of Out of Place subreport

' Vaiables that represent what was passed on the Query String
	Dim CMD					' This determines the main function to perform
	Dim PAGE				' the page to return
	Dim BRCH				' the branch is a mechanism to determine the drill down level.
							' A drill down level is like a view of the report, a new tab
							' is created to indicate that it is a new view
	Dim VIEWER				' This is the viewer that is calling the server
	Dim VFMT				' the format that the viewer understands
	Dim NODE				' Currently not used??
	Dim GRP					' this is a way of specifing the actual group
	Dim COORD				' these are the coordinates on the graph to process
	Dim DIR					' this is the search direction
	Dim CSE					' indicates if the search is case sensitive
	Dim TEXT				' this is the text to search for.
	Dim INIT				' used to build the frames for the html viewer
	Dim NEWBRCH				' used to keep track of when a new branch is to be viewed.
	Dim EXPORT_FMT			' used to hold the export format and type
	Dim SUBRPT				' used to hold the Out Of Place Subreport page, number,
							' and coordinates on the main report.
	Dim INCOMPLETE_PAGE		' used to indicate whether the page generated should contain placeholders.
	Dim INCOMPLETE_PAGE_COUNT ' used to indicate whether the page should contain the total page count if not yet generated.
	Dim PVERSION			' used to indicate the protocol version of the viewer.
	Dim TTL_INFO			' used to indicate the group path of the totaller request.
' Constant Values 
	Dim CREFTWORDFORWINDOWS
	Dim CREFTRICHTEXT
	Dim CREFTEXCEL21
	Dim CREFTCRYSTALREPORT
	Dim CREDTDISKFILE
	Dim EMFMIMETYPE		
	CREFTWORDFORWINDOWS = 14
	CREFTRICHTEXT = 4
	CREFTEXCEL21 = 18
	CREFTCRYSTALREPORT = 1
	CREDTDISKFILE = 1
	crAllowPlaceHolders = 2
	crDelayTotalPageCountCalc = 1
	EMFMIMETYPe = "application/x-emf"
	EPFMIMETYPE = "application/x-epf"
	ETFMIMETYPE = "application/x-etf"
'	Initialize Arrays
	gvGroupPath = Array()
	gvGroupLevel = Array()
	gvMaxNode = Array() ' reteive all nodes
	gvTotallerInfo = Array()
	NEWBRCH	 = "0"
'  To ensure that the browser does not cache the html pages for the group trees.
Response.Expires = 0
' Parse Query String for paramaters
Call ParseQS()

' INIT is a special QS case, we only care about HTML viewer, if it is then save send page and branch info
' to the frame page

if INIT = "HTML_FRAME" then
	' build URL and send the QS
	if BRCH <> ""  and NEWBRCH	= "1" then
		' htmstart is the base page that creates the frames for the HTML viewer
		' if there is branch information it needs to be passed along.
			tmpArray = session("tabArray")
			if tmpArray(0) <> "EMPTY" then
				val = UBound(tmpArray, 1) + 1
				redim preserve tmpArray(val + 4)
			else
				val = 0
			end if
			tmpArray(val) = CStr(val)
			tmpArray(val + 1) = session("lastBrch")
			session("lastBrch") = BRCH
			tmpArray(val + 2) = session("CurrentPageNumber")
			tmpArray(val + 3) = session("lastknownpage")
			tmpArray(val + 4) = session("LastPageNumber") 
			session("tabArray") = tmpArray
		gvURL = "htmstart.asp?brch=" & BRCH & "&"
	else
		if BRCH <> "" then
			gvURL = "htmstart.asp?brch=" & BRCH
		else 
			gvURL = "htmstart.asp"
		end if
	end if
	response.redirect gvURL
end if


	
' If there is a BRCH then create the gvGroupPath array that represents it.

if BRCH <> "" then
	gvGroupPath = CreateArray(BRCH)
end if 

' If there is a GRP then create the gvGroupLevel array that represents it.

if GRP <> "" then
	gvGroupLevel = CreateArray(GRP)
end if

' If there is a TTL_INFO then create the gvTotallerInfo array that represents it.

if TTL_INFO <> "" then
	gvTotallerInfo = CreateArray(TTL_INFO)
end if



' If there are COORDs, then get them
if COORD <> "" then
	Call GetDrillDownCoordinates(COORD, glX, glY)
end if
	
' This case statement determines what action to perform based on CMD
' there are sub cases for each viewer type


		
Select Case CMD

Case "GET_PG"

	Call RetrieveObjects
	
	' create the actual page
	Set goPage = goPageCollection(PAGE)
	' check for an exception on the page number 
	Call ValidatePageNumber

	' 0 is for epf, 8209 is a SafeArray
	Select Case VFMT
	
		Case "ENCP"
			session("oPageEngine").PlaceHolderOptions = 0
			if(PVERSION > 2)then
				if INCOMPLETE_PAGE > 0 then
					session("oPageEngine").PlaceHolderOptions = crAllowPlaceHolders
				end if
				if INCOMPLETE_PAGE_COUNT > 0 then
					session("oPageEngine").PlaceHolderOptions = session("oPageEngine").PlaceHolderOptions + crDelayTotalPageCountCalc
				end if
			end if 				
			session("oPageEngine").ImageOptions = 1
			temp = goPage.Renderepf(8209)
			Response.AddHeader "CONTENT-LENGTH", lenb(temp)
			Response.ContentType = EPFMIMETYPE
			response.binarywrite temp
		
		Case "HTML_FRAME"
			session("oPageEngine").ImageOptions = 1
			response.binarywrite goPage.Renderhtml(1,2,1,request.ServerVariables("SCRIPT_NAME"),8209)
			' Need to know if it is the last page to construct the toolbar correctly
			if goPage.IsLastPage then
				session("LastPageNumber") = goPage.pagenumber
				session("CurrentPageNumber") = session("LastPageNumber")
			end if	
		Case "HTML_PAGE"
			session("oPageEngine").ImageOptions = 1
			response.binarywrite goPage.Renderhtml(1,3,3,request.ServerVariables("SCRIPT_NAME"),8209)
		
		end select
	
Case "GET_TTL"
	
	Call RetrieveObjects
	
	Select Case VFMT
	
		Case "ENCP"
			Response.ContentType = ETFMIMETYPE
			if(PVERSION > 2)then
				temp = goPageGenerator.RenderTotallerETF(gvTotallerInfo, 0, 1, gvMaxNode, 8209)
			else
				temp = goPageGenerator.RenderTotallerETF(gvGroupPath, 0, 0, gvMaxNode, 8209)
			end if
			Response.AddHeader "CONTENT-LENGTH", lenb(temp)
			response.binarywrite temp
		
		Case "HTML_FRAME"
			response.binarywrite goPageGenerator.RenderTotallerHTML(gvGroupPath, 1, 0, gvMaxNode, gvGroupLevel, 1, request.ServerVariables("SCRIPT_NAME"), 8209)
		
	end select


Case "RFSH"
	
	' This command forces the database to be read again.
	session("oRpt").DiscardSavedData
	session("oRpt").ReadRecords
	If Err.Number <> 0 Then
		Call CheckForError
	Else
		session("oRpt").EnableParameterPrompting = False
		Set session("oPageEngine") = session("oRpt").PageEngine
	End If
	Call RetrieveObjects
	Set goPage = goPageCollection(PAGE)
	Call ValidatePageNumber
	session("oPageEngine").ImageOptions = 1
	Select Case VFMT
	Case "ENCP"
	' Java and Active X Viewers will make a get page command when receiving 0 error msg value
		if VIEWER = "JAVA" then
			session("oPageEngine").PlaceHolderOptions = 0
			if(PVERSION > 2)then
				if INCOMPLETE_PAGE > 0 then
					session("oPageEngine").PlaceHolderOptions = crAllowPlaceHolders
				end if
				if INCOMPLETE_PAGE_COUNT > 0 then
					session("oPageEngine").PlaceHolderOptions = session("oPageEngine").PlaceHolderOptions + crDelayTotalPageCountCalc
				end if
			end if 
			temp = goPage.Renderepf(8209)
			Response.AddHeader "CONTENT-LENGTH", lenb(temp)
			Response.ContentType = EPFMIMETYPE
			response.binarywrite temp
		else
			Response.ContentType = EMFMIMETYPE
			session("oEMF").SendErrorMsg 0,""
		end if 
		
	Case "HTML_FRAME"
		InitializeFrameArray()
		gvURL = "htmstart.asp"
		response.redirect gvURL

	Case "HTML_PAGE"
	session("oPageEngine").ImageOptions = 1
	response.binarywrite goPage.Renderhtml(1,3,1,request.ServerVariables("SCRIPT_NAME"),8209)
		
	end select


Case "NAV"
	Call RetrieveObjects
	Call CheckForError
	' Get the page number that the group in on, for this particular branch
	gvPageNumber = goPageGenerator.GetPageNumberForGroup(gvGroupLevel)
			
	Select Case VFMT
	' 0 is for epf, 8209 is a SafeArray, 8 is a BSTR
	Case "ENCP"
		' Create a byte array for the EMF, which will contain the page number
		Response.ContentType = EMFMIMETYPE
		session("oEMF").sendpagenumberrecord(gvPageNumber)
	
	Case "HTML_FRAME"
		' for html browser send back the page
		dim appendQuery
		appendQuery = "?"
		session("CurrentPageNumber") = gvPageNumber
		if BRCH <> "" then
			appendQuery = appendQuery & "BRCH=" & BRCH & "&"
		end if
		if GRP <> "" then
				appendQuery = appendQuery & "GRP=" & GRP
		end if
		response.redirect "framepage.asp" & appendQuery
	
	end select


Case "CHRT_DD"
	' only supported in java and active X smart viewers
	Select Case VFMT

	Case "ENCP"

		'  Get page collection
		Call RetrieveObjects
		Call CheckForError
		' Pass the coordinates to the report engine to determine what
		' branch the drill down goes to.
		Set goPageGeneratorDrill = goPageGenerator.DrillOnGraph(PAGE, glX, glY)
		' Check for an exception because of coordinates
		if err.number <> 0 then
			gsErrorText = "Not part of the Graph "
			Response.ContentType = EMFMIMETYPE
			session("oEMF").SendErrorMsg 40, gsErrorText
			err.clear
			response.end
		end if
		' pass the group level and group path to helper function to create 
		' the EMF message, this tells the viewer where to get the page.

		gvGroupPathDD = goPageGeneratorDrill.grouppath
		gvGroupNameDD = goPageGeneratorDrill.groupname
		Response.ContentType = EMFMIMETYPE
		session("oEMF").GroupName = gvGroupNameDD		
		session("oEMF").sendbranchesemf(gvGroupPathDD)		
			

	end select

Case "GET_LPG"
	
	' only support in smart viewers
	Select Case VFMT

	Case "ENCP"
		' this command returns the page number of the last page
		' Get page collection
		Call RetrieveObjects
		Call CheckForError
		' Get the count from the Pages collection
		gvPageNumber = goPageCollection.Count

		' Send the EMF representing the page number
		Response.ContentType = EMFMIMETYPE
		session("oEMF").sendpagenumberrecord(gvPageNumber)
	end select

Case "SRCH"
	Call RetrieveObjects
	Call CheckForError
	' create page variable
	gvPageNumber = CInt(PAGE)
	
	Select Case VFMT
	Case "ENCP"
		if goPageGenerator.FindText(TEXT, 0, gvPageNumber) then
			Response.ContentType = EMFMIMETYPE
			session("oEMF").sendpagenumberrecord(gvPageNumber)
		else
			gsErrorText = "The specified text, '" & TEXT & "' was not found in the report"
			Response.ContentType = EMFMIMETYPE
			session("oEMF").SendErrorMsg 33, gsErrorText
		end if
					
	Case "HTML_FRAME"
		' We are being called by HTML viewer
		' need to get the text from the form post
		dim searchFound
		TEXT = request.form("text")
		' Now find out what page the text is on
		tempNumber = gvPageNumber + 1
		If(CBool(goPageGenerator.FindText(TEXT, 0, tempNumber))) then
			session("CurrentPageNumber") = tempNumber
			searchFound = 1
		else
			session("CurrentPageNumber") = gvPageNumber
			searchFound = 0
		End If
		if BRCH <> "" then
			gvURL = "framepage.asp?brch=" & BRCH & "&SEARCHFOUND=" & searchFound
		else
			gvURL = "framepage.asp?SEARCHFOUND=" & searchFound
		end if
		response.redirect gvURL

	Case "HTML_PAGE"
		' We are being called by HTML viewer
		' need to get the text from the form post
		TEXT = request.form("text")
		' Now find out what page the text is on
		tempNumber = gvPageNumber
		If(CBool(goPageGenerator.FindText(TEXT, 0, tempNumber))) then
			gvPageNumber = tempNumber
			Set goPage = goPageCollection(gvPageNumber)
			session("oPageEngine").ImageOptions = 1
			response.binarywrite goPage.Renderhtml(1,3,3,request.ServerVariables("SCRIPT_NAME"),8209)
		else
		' Send back an html page indicating the text was not found.
			Response.Write "<html><title>Seagate ASP Reports Server</title><body bgcolor='white'><center><h1>The text cannot be found in this report.</h1></center></body></html>"
		End If
		
	end select

				
Case "TOOLBAR_PAGE"
	
	' Redirect to the framepage, need to know if we are 
	' on the last page.

	if session("LastPageNumber") <> "" then
		if CInt(PAGE) > CInt(session("LastPageNumber")) then
			session("CurrentPageNumber") = session("LastPageNumber")
		else
			session("CurrentPageNumber") = PAGE
		end if
	else 
		Call RetrieveObjects
		Call CheckForError
		' create the actual page
		Set goPage = goPageCollection(PAGE)
		' check for an exception on the page number 
		Call ValidatePageNumber
		if goPage.IsLastPage then
			session("LastPageNumber") = goPage.pagenumber
			session("CurrentPageNumber") = session("LastPageNumber")	
		else
			session("CurrentPageNumber") = PAGE
		end if	
	end if
	if BRCH <> "" then
		gvURL = "framepage.asp?brch=" & BRCH
	else
		gvURL = "framepage.asp"
	end if

	response.redirect gvURL

Case "EXPORT"
	Set ExportOptions = Session("oRpt").ExportOptions
	if(FillExportOptionsObject( EXPORT_FMT)) Then
		Call RetrieveObjects
		response.binarywrite goPageGenerator.Export(8209)
		Call CheckForError
	else
		Response.ContentType = EMFMIMETYPE
		session("oEMF").SendErrorMsg 1, "Invalid Export Type Specified"
	end if

Case "MAP_DD"
	' only supported in java and active X smart viewers
	Select Case VFMT

	Case "ENCP"

		'  Get page collection
		Call RetrieveObjects
		Call CheckForError
		' Pass the coordinates to the report engine to determine what
		' branch the drill down goes to.
		Set goPageGeneratorDrillonMap = goPageGenerator.DrillOnMap(PAGE, glX, glY)
		' Check for an exception because of coordinates
		if err.number <> 0 then
			gsErrorText = "No Values Exist for Selected Region of Map"
			Response.ContentType = EMFMIMETYPE
			session("oEMF").SendErrorMsg 40, gsErrorText		
			err.clear
			response.end
		end if
		' pass the group level and group path to helper function to create 
		' the EMF message, this tells the viewer where to get the page.

		gvGroupPathDD = goPageGeneratorDrillonMap.grouppath
		gvGroupNameDD = goPageGeneratorDrillonMap.groupname
		session("oEMF").GroupName = gvGroupNameDD	
		Response.ContentType = EMFMIMETYPE	
		session("oEMF").sendbranchesemf(gvGroupPathDD)		
			
	end select

end select



SUB RetrieveObjects() 
' This procedure simply retrieves the session objects into global variables.
' In the case of Out of Place Subreports, the SUBRPT parameter must be parsed and the
' Subreport page generator object must be created.
	Dim oRptOptions 'Report Options 
	Dim charIndexVal,tmpCharIndexVal
	Dim tmpStr
	Dim tmpPageGenerator
	Dim subPageGenerator 
	Dim OOPSSeqNo	'holds the page's OOPS sequence number
	Dim OOPSSubName	'holds the OOPS's name
	Dim subCoords 'holds the coordinates of the OOPS in the main report
	Dim subgvGroupPath 'holds the group path for the main report in subrpt parameter
	Dim mainRptPageNumber 'holds the page number for the main report in the subrpt parameter
	
	subgvGroupPath = Array()
	if IsObject(session("oPageEngine")) then
		' make sure dialogs have been disabled
		if SUBRPT <> "" Then
		' Obtain the subreport sequence number
			charIndexVal = findChar(SUBRPT, ":")
			if charIndexVal > 1 then
				OOPSSeqNo = Mid(SUBRPT,1,charIndexVal - 1)
			end if
		' Obtain the subreport's name
			tmpStr = Mid(SUBRPT,charIndexVal + 1)
			charIndexVal = findChar(tmpStr, ":")
			if charIndexVal > 1 then
				OOPSSubName = Mid(tmpStr,1,charIndexVal - 1)
			end if
			tmpStr = Mid(tmpStr,charIndexVal + 1)
			charIndexVal = findChar(tmpStr, ":")
		' Obtain the group path for the Out of Place Subreport
			if charIndexVal > 1 then
				subgvGroupPath = CreateArray(Mid(tmpStr, 1, charIndexVal - 1))
			end if
		'Obtain the main report page number after the fourth : character
			tmpStr = Mid(tmpStr,charIndexVal + 1)
		'Get the location of the fourth : seperator
			charIndexVal = findChar(tmpStr, ":")
			mainRptPageNumber = Mid(tmpStr, 1, charIndexVal - 1)
		'Get the coordinates portion of the SUBRPT parameter
			subCoords = Mid(tmpStr, charIndexVal + 1)
			Call GetDrillDownCoordinates(subCoords, slX, slY)
			' Get the main reports page generator for the view
			Set tmpPageGenerator = session("oPageEngine").CreatePageGenerator(subgvGroupPath)
			Set subPageGenerator = tmpPageGenerator.DrillOnSubreport(mainRptPageNumber, slX, slY)
			Set goPageGenerator = subPageGenerator.CreateSubreportPageGenerator(gvGroupPath)
		else
			Set goPageGenerator = session("oPageEngine").CreatePageGenerator(gvGroupPath)
			end if
		Set goPageCollection = goPageGenerator.Pages
	else
		' must have timed out return an error, you may wan to Append to the
		' IIS log here.
		if VFMT = "ENCP" then 
			Response.ContentType = EMFMIMETYPE
			session("oEMF").SendErrorMsg 1, "User Session has expired"
		else
			response.write "User Session has expired"
			
		end if
		response.end
	end if

END SUB

SUB ParseQS()
	' Parse the Query String 
	CMD = UCase(request.querystring("cmd"))		' This determines the main function to perform
	PAGE = UCase(request.querystring("page"))	' the page to return
	BRCH = UCase(request.querystring("BRCH"))	' the branch is a mechanism to determine the drill down level.
												' A drill down level is like a view of the report, a new tab
												' is created to indicate that it is a new view
	VIEWER = UCase(request.querystring("VIEWER"))	' This is the viewer that is calling the server
	VFMT = UCase(request.querystring("VFMT"))	' the format that the viewer understands
	NODE = UCase(request.querystring("NODE"))
	GRP = UCase(request.querystring("GRP"))		' this is a way of specifing the actual group
	COORD = UCase(request.querystring("COORD"))	' these are the coordinates on the graph to process
	DIR = UCase(request.querystring("DIR"))		' this is the search direction
	CSE = UCase(request.querystring("CASE"))	' indicates if the search is case sensitive
	TEXT = request.querystring("TEXT")			' this is the text to search for.
	INIT = UCase(request.querystring("INIT"))	' used to build the frames for the html viewer
	TAB = UCase(request.querystring("TAB"))		' used to keep track of TABS on drill down.
	EXPORT_FMT = UCase(request.querystring("EXPORT_FMT")) ' Used to specify export format and type.	
	SUBRPT = UCase(request.querystring("SUBRPT")) ' The Out of Place Subreport coordinates.
	INCOMPLETE_PAGE = CInt(request.querystring("INCOMPLETE_PAGE"))' Used to specify whether the page is to contain placeholders.
	INCOMPLETE_PAGE_COUNT = CInt(request.querystring("INCOMPLETE_PAGE_COUNT"))' Used to specify whether the page has to contain a total page count.
	PVERSION = CInt(request.querystring("PVERSION"))' Used to indicate the protocol version the viewer is utilizing.
	TTL_INFO = UCase(request.querystring("TTL_INFO"))'Used to indicate the group path of the totaller request.
	' Initialize variables to a default if they are not provided on the query string.
	' Check for Parameter Values that are passed by the HTTP Post Command.
	if CMD = "" then
		CMD = UCase(request.form("cmd"))	
		if CMD = "" then
			CMD = "GET_PG"
		end if
	end if
	
	if INIT = "" then
		INIT = UCase(request.form("INIT"))
	end if

	if BRCH = "" then
		BRCH = UCase(request.form("BRCH"))
	end if

	if BRCH = "" and INIT = "HTML_FRAME" then 
		Call InitializeFrameArray
	end if


	if BRCH <> "" and INIT = "HTML_FRAME"  then
		if session("lastBrch") <> BRCH then
			NEWBRCH	 = "1"
		end if
	end if 
			

	if VIEWER = "" then
		VIEWER = UCase(request.form("VIEWER"))
		if VIEWER = "" then
			VIEWER = "HTML"
		end if
	end if

	if VFMT = "" then 
		VFMT = UCase(request.form("VFMT"))
		if VFMT = "" then 
			VFMT = "HTML_PAGE"
		end if
	end if

	if GRP = "" then
		GRP = UCase(request.form("GRP"))	
	end if

	if TTL_INFO = "" then
		TTL_INFO = UCase(request.form("TTL_INFO"))
	end if

	if COORD = "" then
		COORD = UCase(request.form("COORD"))
	end if

	if NODE = "" then
		NODE = UCase(request.form("NODE"))
	end if

	if DIR = "" then
		DIR = UCase(request.form("DIR"))
		if DIR = "" then
			DIR = "FOR" ' forward
		end if
	End if

	if CSE = "" then
		CSE = UCase(request.form("CASE"))
		if CSE = "" then
			CSE = "0" ' case insensitive
		end if
	end if

	if TEXT = "" then
		TEXT = request.form("TEXT")
	end if

	if EXPORT_FMT = "" then
		EXPORT_FMT = UCase(request.form("EXPORT_FMT"))
	end if
	
	if SUBRPT = "" then
		SUBRPT = UCase(request.form("SUBRPT"))
	end if
	
	if request.form("INCOMPLETE_PAGE") <> "" then
		INCOMPLETE_PAGE = CInt(request.form("INCOMPLETE_PAGE"))
	end if
	
	if request.form("INCOMPLETE_PAGE_COUNT") <> "" then
		INCOMPLETE_PAGE_COUNT = CInt(request.form("INCOMPLETE_PAGE_COUNT"))
	end if
	
	if PVERSION = 0 then
		PVERSION = CInt(request.form("PVERSION"))
	end if

' Check to make sure there is a page requested, if not use 1 as a default
	if PAGE = "" then
		PAGE = UCase(request.form("page"))
		if PAGE = "" then
			PAGE = "1"
		end if
	end if
	
	if PAGE <> "" and NOT IsNumeric(PAGE) then
		PAGE = "1"
	end if
	
END SUB

Function CreateArray(ByVal vsStringArray)
' this function takes an string like 0-1-1-0 and converts
' it into an array of integers

    Dim lvArray
    Dim lvNewArray
    Dim liCount
    Dim liCurrentPos
    Dim lsBuf
    lvArray = Array()
    lvNewArray = Array()
    ReDim lvArray(256)
    
    liStringLength = Len(vsStringArray)
    liCount = 0
    liCurrentPos = 1
    lsBuf = ""
    
    While liCurrentPos <= liStringLength
         
         'ignore this character
        If Mid(vsStringArray, liCurrentPos, 1) <> "-" Then
            lsBuf = lsBuf & Mid(vsStringArray, liCurrentPos, 1)
            If liCurrentPos = liStringLength Then
                lvArray(liCount) = CInt(lsBuf)
                lsBuf = ""
                liCount = liCount + 1
            End If
            
        Else
            lvArray(liCount) = CInt(lsBuf)
            lsBuf = ""
            liCount = liCount + 1
        End If
        liCurrentPos = liCurrentPos + 1
    Wend
    
    ReDim lvNewArray(liCount - 1)
    
    For x = 0 To (liCount - 1)
        lvNewArray(x) = lvArray(x)
    Next
    
    
    CreateArray = lvNewArray

End Function

' Helper function to parse coordinates passed by viewers and place into independent variables.
SUB GetDrillDownCoordinates(ByVal strParam, ByRef xCoord, ByRef yCoord)
	Dim liStringLength
	Dim lbDone
	Dim lsBuf

	liStringLength = Len(strParam)
	lbDone = FALSE
	lsBuf = ""
	xCoord = ""
	yCoord = ""
	For x = 1 To liStringLength
		lsBuf = Mid(strParam, x, 1)
		
		'ignore this character
		If lsBuf = "-" Then
			lsBuf = ""
			lbDone = TRUE
		End if
		
		if lbDone then
			yCoord = yCoord + lsBuf
		else
			xCoord = xCoord + lsBuf
		end if
			
	Next
	
END SUB

' This helper procedure will check if the requested page number exists.
' If it does not, it will set it to the last available page.
SUB ValidatePageNumber()
	if err.number <> 0 then
		if err.number = 9 then 
			' just return the last page
			PAGE = goPageCollection.count
			Set goPage = goPageCollection(PAGE)
			' these session variables are used for the HTML Frame viewer
			session("LastPageNumber") = PAGE
			session("CurrentPageNumber") = PAGE
			err.clear
		else
			' A more serious error has occurred. Error message sent to browser.
			Call CheckForError
		end if
	end if
END SUB



'  This helper procedure will send an error msg to the browser based on what viewer is being used.
SUB CheckForError()
	If Err.Number <> 0 Then
		if VFMT = "ENCP" then
			Response.ContentType = EMFMIMETYPE
			session("oEMF").SendErrorMsg 1, "CRAXDRT Error Occured on Server. " & Err.Number  & " : " & Err.Description
		else
			Response.Write "CRAXDRT Error Occured on Server. Error Number: " & Err.Number & " Error Description: " & Err.Description
		end if
		Response.end
	End if
END SUB

SUB InitializeFrameArray()
'initialize the html_frame array
	set session("tabArray") = Nothing
	session("lastBrch") = ""
	dim tmpArray
	tmpArray = Array(4)
	redim tmpArray(4)
	'Initialize the sequence number
	tmpArray(0) = "EMPTY"
	session("tabArray") = tmpArray
END SUB

' Helper function to parse the EXPORT_FMT parameter and fill in the properties of the 
' Export object.
FUNCTION FillExportOptionsObject(export_fmt_options)
	dim charIndex 
	dim exportType
	dim exportDLLName
	charIndex = findChar(export_fmt_options,":")
	'Set session("ExportOptions") = Session("oRpt").ExportOptions
	if(charIndex > 0) Then
	'Get the export format type value
		exportType = Mid(export_fmt_options, charIndex + 1)
		exportDLLName = UCase(Mid(export_fmt_options, 1, charIndex - 1))
		Select Case exportDLLName
			Case "U2FWORDW"
				ExportOptions.FormatType = 	CREFTWORDFORWINDOWS + CInt(exportType)
				Response.ContentType = "application/msword"
			Case "U2FRTF"
				ExportOptions.FormatType = 	CREFTRICHTEXT	+ CInt(exportType)
				Response.ContentType = "application/rtf"
			Case "U2FXLS"
				ExportOptions.FormatType = CREFTEXCEL21	+ CInt(exportType)
				Response.ContentType = "application/x-msexcel"
			Case "U2FCR"
				ExportOptions.FormatType = CREFTCRYSTALREPORT	+ CInt(exportType)
				Response.ContentType = "application/x-rpt"
			Case Else
				FillExportOptionsObject = False
				Exit Function
		End Select
		ExportOptions.DestinationType = CREDTDISKFILE
		FillExportOptionsObject = True
	else
		FillExportOptionsObject = False
	end if

end FUNCTION
		 
'  Helper function that returns the index of the character in the given string.
Function findChar(findStr, charToFind)
	dim lenStr 
	dim result 
	lenStr = len(findStr)
	result = -1
	if(lenStr > 0) Then
		charCounter = 1
		do While(charCounter <= lenStr)
			tmpChar = Mid(findStr,charCounter,1)
			if(tmpChar = charToFind) Then
				result = charCounter
				Exit Do
			end if
			charCounter = charCounter + 1
		loop
	end if
	
	findChar = result
End Function	

%>


