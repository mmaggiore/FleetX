<%
'''''''''''''''''''''''''''''''
'''''''TEST VS. PRODUCTION
	TaskmasterDirectory="Intranet/FleetX/login.asp"
	something=Request.ServerVariables("HTTP_HOST") 
	'response.write "something="&something&"<BR>"
	somethingelse=Request.ServerVariables("URL") 
	'response.write "somethingelse="&somethingelse&"<BR>"



	ServerType=lcase(left(Something,4))
	Select Case ServerType
		Case "test"
			WhichSite="https://test.logisticorp.us/Intranet/fleetx"
			DATABASE="DATABASE=FleetX;DSN=SQLConnect;UID=sa;Password=cadre;"
			INTRANET="DATABASE=lcintranet;DSN=Intranet;UID=sa;Password=cadre;"
			SiteName="TEST"
			ShowSiteName="y"
		Case "demo"
			WhichSite="https://demo.logisticorp.us/Intranet/fleetx"
			DATABASE="DATABASE=FleetXDemo;DSN=SQLConnect;UID=sa;Password=cadre;"
			INTRANET="DATABASE=lcintranetDemo;DSN=Intranet;UID=sa;Password=cadre;"
			SiteName="DEMO"
			ShowSiteName="y"
		Case else
			'If lcase(left(Somethingelse,4))="/int" then
				WhichSite="https://www.logisticorp.us/fleetx"
				'else	
				'WhichSite="https://www.logisticorp.us/fleetX"
			'End if
			DATABASE="DATABASE=FleetX;DSN=SQLConnect;UID=sa;Password=cadre;"
			INTRANET="DATABASE=lcintranet;DSN=Intranet;UID=sa;Password=cadre;"
			SiteName="PRODUCTION"
			ShowSiteName="n"

	End Select
	'''''''''''''''''''''''''''''''''''''''''''''
	''''''''CRYSTAL REPORTS DATABASE INFO''''''''
	Session("ReportAppServerVariable")="192.168.104.230"
	'''''''''''''''''''''''''''''''''''''''''''''
	'Response.write "WhichSite="&WhichSite&"<BR>"



'============================================================================
'				  Important DATABASE RELATED Constants  
' System Settings
sBT_ID=86

''''''''''''''''''''''Here's new one for Liberty IMS
LIBERTY="DATABASE=LogistiCorpData;DSN=LibertyIMS;UID=sa;Password=cadre;"
'LIBERTYIMS="DATABASE=LogistiCorpData;DSN=LibertyIMS;UID=sa;Password=cadre;"
'INTRANET="DATABASE=lcintranet;DSN=Intranet;UID=sa;Password=cadre;"
'InventoryDB="DATABASE=Inventory;DSN=Intranet;UID=sa;Password=cadre;"

CopywriteNotice="&#169; LogistiCorp Group LLC.  All Rights Reserved."


URL=(Request.ServerVariables("url"))
If lcase(right(URL, 9))="login.asp" or lcase(right(URL, 9))="?logout=y" or lcase(right(URL, 11))="newuser.asp" or lcase(right(URL, 19))="requestuserinfo.asp" or lcase(right(URL, 18))="nitelydispatch.asp" then
	'Response.write "LOGIN11111<BR>"

	Else
	'Response.write "VERIFY IF LOGGED IN!!!!<BR>"
	UserID=Request.cookies("FleetXCookie")("UserID")
	'Response.write "UserID="&UserID&"<BR>"
	UserName=Request.cookies("FleetXCookie")("UserName")
	'Response.write "UserID="&UserID&"<BR>"
	RequestorCompany=Request.cookies("FleetXCookie")("RequestorCompany")
	'Response.write "RequestorCompany="&RequestorCompany&"<BR>"
	BillToID=Request.cookies("FleetXCookie")("BillToID")
	'Response.write "BillToID="&BillToID&"<BR>"
	'Userid=""
	If trim(UserID)="" then UserID=Session("UserID") end if
	If trim(RequestorCompany)="" then RequestorCompany=Session("RequestorCompany") end if
	If trim(UserName)="" then UserName=Session("UserName") end if
	If trim(BillToID)="" then BillToID=Session("BillToID") end if
	'If Trim(UserID)="" then 
	'	response.redirect(WhichSite&"/Login.asp") 
	'	'response.write "GOT HERE!!!!!<BR>"
	'end if
	''''''Resets cookies to extend session.
	Response.Cookies ("FleetXCookie")("RequestorCompany")=RequestorCompany
	Response.Cookies("FleetXCookie").expires = dateadd("n",+120,now())
    Response.Cookies ("FleetXCookie")("UserID") = UserID
    Response.Cookies("FleetXCookie").expires = dateadd("n",+120,now())
    Response.Cookies ("FleetXCookie")("UserName") = UserName
    Response.Cookies("FleetXCookie").expires = dateadd("n",+120,now())
    Response.Cookies ("FleetXCookie")("BillToID") = BillToID
    Response.Cookies("FleetXCookie").expires = dateadd("n",+120,now())


End if



'******************************************************************
'Sanitize request variables for sql injection and malicious code
'******************************************************************
function valid8(aString)

	'Declare Variables
	dim tempString
	tempString = trim(aString)
	
	'Check for empty values
	if isNull(tempString) or isEmpty(tempString) or len(tempString) = 0 then
		'valid8 = ""
		exit function
	end if

	'Sanitize
	tempString = replace(tempString,"<", " ")
	tempString = replace(tempString,">", " ")
	tempString = replace(tempString,"'","''")
	tempString = replace(tempString,"""","'")
	tempString = replace(tempString,"%"," ")
	tempString = replace(tempString,"&"," and ")
	tempString = replace(tempString,","," ")
	tempString = replace(tempString,"(", " ")
	tempString = replace(tempString,")", " ")

	tempString = replace(tempString,"--"," ")
	tempString = replace(tempString,"="," ")
	tempString = replace(tempString,";", " ")
	tempString = replace(tempString,"@@"," ")
	tempString = replace(tempString,"|"," ")

	'tempString = replace(tempString,"cast"," ")
	tempString = replace(tempString,"cast("," ")
	tempString = replace(tempString," cast "," ")
	tempString = replace(tempString,"drop "," ")
	tempString = replace(tempString,"declare"," ")
	'tempString = replace(tempString,"char"," ")
	tempString = replace(tempString,"exec"," ")
	tempString = replace(tempString,"char("," ")
	tempString = replace(tempString," char "," ")
	tempString = replace(tempString,"cursor"," ")
	tempString = replace(tempString,"iframe"," ")

	valid8  = trim(tempString)
	
end function


'******************************************************************
'Convert date to days, hours, mins
'******************************************************************
Function datediffCNV(d1, d2) 
        report = ""
        'Start with total number of days
        days = DateDiff("d",d1,d2)
  		'Convert days to years and grab remaining days
		'If days > 365 Then
			'years = days\365
			'days = days Mod (365*years)-1
			'report = years & " Year(s), "
		'Else 
			'years = 0	
		'End If
		'Thank you PHV for help simplifying the month calculation
		'Compute the number of months
		'months = Int(DateDiff("m",d1,d2))+(day(d2)<day(d1))
		'remove years from the total months
		'months = months Mod 12
		'if months > 0 then
      'report = report & Months & " Month(s), "
    'end if
		
		'now find the days
		'newStart = Month(d1) & "/" & Day(d1) & "/" & Year(d1) + years
		'If Month(d1) <> 12 Then
			'fullmonthStart = Month(d1) + 1 & "/1/" & Year(d1) + years
		'Else
			'fullmonthStart = "1/1/" & Year(d1) + years +1
		'End If
		
		'If Day(d1) =< Day(d2) Then
			'days = Day(d2) - Day(d1)
		'Else
			'days = DateDiff("d", newStart, fullmonthStart) + Day(d2) -1
		'End If
        'If days > 0 Then
        	'report = report & days & " day(s) "
        'End If	
        
        'now we will deal with the time left over
        'begin by getting total seconds between dates and divide out the days
        'grab the remaining seconds with the mod operator
        Seconds = abs(datediff("S", d1, d2)) 
        if Seconds <= 0 then 
            'report = "0 seconds." 
        else 
	        'Seconds = Seconds mod (24*60*60) 
	        'divide by 3600 to get hours
	        If Seconds >= 3600 then 
	            report = report & _  
	            Seconds\(3600) & " hr(s) " 
	        end If
	        'use mod to get remaining seconds and divide to get minutes
	        Seconds = Seconds mod (60*60) 
	        if Seconds >= 60 then 
	            report = report & _  
	            Seconds\(60) & " min(s) " 
	        end If
	        'use mod to get remaining seconds
	        'seconds = Seconds Mod (60) 
	        'report = report & " " & seconds & " sec(s)"
        end if 
        datediffCNV = report 
End Function 


If ShowSiteName="y" then
%>
<table width="100%" bgcolor="black" class="hide">
	<tr><td align="center"><font color="white"><b><%=SiteName%></b></font></td></tr>
</table>
<%
End if
%>

