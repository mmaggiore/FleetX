<%
Dim iMsg, iConf, Flds
'''''''''''''''''''''''''''''''
'''''''TEST VS. PRODUCTION
	TaskmasterDirectory="driverlogin.asp"
'''''''TEST VS. PRODUCTION
	TaskmasterDirectory="Intranet/FleetX/login.asp"
	something=Request.ServerVariables("HTTP_HOST") 
	'response.write "something="&something&"<BR>"
	somethingelse=Request.ServerVariables("URL") 
	'response.write "somethingelse="&somethingelse&"<BR>"


'''''''''''''''''''''''''''''''''''''''''''




	ServerType=lcase(left(Something,4))
	Select Case ServerType
		Case "10.1" 'TEST
			WhichSite="HTTP://10.10.51.102/fleetx"
			varReports="HTTP://10.10.51.121"
			DATABASE="DATABASE=FleetX;DSN=SQLConnect;"
			INTRANET="DATABASE=lcintranet;DSN=Intranet"
			SiteName="TEST"
			ShowSiteName="y"
			Response.write"<table width='100%'><tr><td bgcolor='red' align='center'><B>TEST</B></td></tr></table>"
		Case "demo" 'DEMO
			WhichSite="HTTPS://demo.logisticorpgroup.com/fleetx"
			varReports="HTTPS://10.10.10.39"
			DATABASE="DATABASE=FleetXDemo;DSN=SQLConnect"
			INTRANET="DATABASE=lcintranetDemo;DSN=Intranet"
			SiteName="DEMO"
			ShowSiteName="y"
			Response.write"<table width='100%'><tr><td bgcolor='blue' align='center'><B>DEMO</B></td></tr></table>"
		Case else
			'If lcase(left(Somethingelse,4))="/int" then
				WhichSite="HTTPS://apps.logisticorpgroup.com/fleetx"
				varReports="HTTPS://10.10.10.121"
				'else	
				'WhichSite="https://www.logisticorp.us/fleetX"
			'End if
			Response.write"<table width='100%'><tr><td bgcolor='green' align='center'><B>PRODUCTION</B></td></tr></table>"
			'Commented out Jason Freis
			'DATABASE="DATABASE=FleetX;DSN=SQLConnect;UID=sa;Password=cadre;"
			'INTRANET="DATABASE=lcintranet;DSN=Intranet;UID=sa;Password=cadre;"
			DATABASE="DATABASE=FleetX;DSN=SQLConnect;"
			INTRANET="DATABASE=lcintranet;DSN=Intranet;"
			SiteName="PRODUCTION"
			ShowSiteName="n"

	End Select
'Jason Freis - WorkflowStudios
'01/22/2017 Added AWS Email CONSTANTS


Const AWS_SendUsingPort = 2
Const AWS_SMTPServer = "email-smtp.us-west-2.amazonaws.com"
Const AWS_SMTPUseSSL = True
'Const AWS_SMTPServerPort = 25
Const AWS_SMTPServerPort = 465
Const AWS_SMTPAuthenticate = 1
Const AWS_SendUserName = "AKIAI7T655UKKIZXNVLA"
Const AWS_SendPassword = "AuP8cMEst6iFplnlYJjrtvfP4nfRCDfjhc/B4CQq2yx0"
Const AWS_SMTPConnectionTimeout = 10
'-----------------------------------------------------------------------------
'============================================================================
'				  Important DATABASE RELATED Constants  
' System Settings
'''sBT_ID=86
'''DATABASE="DATABASE=FleetX;DSN=SQLConnect;UID=sa;Password=cadre;"
'''INTRANET="DATABASE=lcintranet;DSN=Intranet;UID=sa;Password=cadre;"
''''''''''''''''''''''Here's new one for Liberty IMS
'''LIBERTYIMS="DATABASE=LogistiCorpData;DSN=LibertyIMS;UID=sa;Password=cadre;"
'''INTRANET="DATABASE=lcintranet;DSN=Intranet;UID=sa;Password=cadre;"
'''InventoryDB="DATABASE=Inventory;DSN=Intranet;UID=sa;Password=cadre;"

CopywriteNotice="&#169; LogistiCorp Group LLC.  All Rights Reserved."


URL=(Request.ServerVariables("url"))
If lcase(right(URL, 9))="login.asp" or lcase(right(URL, 9))="?logout=y" or lcase(right(URL, 11))="newuser.asp" or lcase(right(URL, 19))="requestuserinfo.asp" then
	'Response.write "LOGIN11111<BR>"

	Else
	'Response.write "VERIFY IF LOGGED IN!!!!<BR>"
	UserID=Request.cookies("FleetXPhone")("UserID")
	'Response.write "UserID="&UserID&"<BR>"
	DriverFirstName=Request.cookies("FleetXPhone")("DriverFirstName")
	DriverLastName=Request.cookies("FleetXPhone")("DriverLastName")
	VehicleType=Request.cookies("FleetXPhone")("VehicleType")
	'Response.write "UserID="&UserID&"<BR>"
	'Response.write "DriverFirstName="&DriverFirstName&"<BR>"
	'Response.write "DriverLastName="&DriverLastName&"<BR>"
	'Userid=""
	If Trim(UserID)="" then 
		'Response.write "FleetX.inc line 42<BR>"
		response.redirect("DriverLogin.asp") 
	end if
	''''''Resets cookies to extend session.
	'Response.write "Reset my cookies!  "& now() &"<BR>"
    Response.Cookies ("FleetXPhone")("UserID") = UserID
    Response.Cookies("FleetXPhone").expires = dateadd("n",+120,now())
    Response.Cookies ("FleetXPhone")("DriverFirstName") = DriverFirstName
    Response.Cookies("FleetXPhone").expires = dateadd("n",+120,now())
    Response.Cookies ("FleetXPhone")("DriverLastName") = DriverLastName
    Response.Cookies("FleetXPhone").expires = dateadd("n",+120,now())

End if
'''''''NEW CODE TO ELIMINATE ACCIDENTAL LOGOUTS!
VehicleID=Request.Cookies("FleetXPhone")("VehicleID")
VehicleName=Request.Cookies("FleetXPhone")("VehicleName")
VehicleType=Request.Cookies("FleetXPhone")("VehicleType")


url = Request.ServerVariables ("URL")
arrayX = split(url,"/")
PageName=lcase(ArrayX(ubound(arrayX)))
'Response.write "PageName="&PageName&"<BR>"
'''''''''''''''''''''''''''''''''''''''''''''''''
'Response.write "UserID="&UserID&"<BR>"
'Response.write "VehicleID="&VehicleID&"<BR>"
'Response.write "LoginCheck="&LoginCheck&"<BR>"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If UserID>"" and VehicleID>"" and LoginCheck<>"n" then
	'Response.write "I GOT HERE!  BOOYAH!<BR>"
			Set Recordset2 = Server.CreateObject("ADODB.Recordset")
			Recordset2.ActiveConnection = Database
			Recordset2.Source = "SELECT VehicleID AS TempVehicleID, DriverID as TempDriverID FROM"
            Recordset2.Source = Recordset2.Source&" AvailableVehicles"
            Recordset2.Source = Recordset2.Source&" WHERE (DriverID='"&trim(UserID)&"') AND (VehicleID='"&trim(VehicleID)&"') and AvailableStatus='c' "
			'Response.Write "SQL="&Recordset2.Source&"<BR>"
			Recordset2.CursorType = 0
			Recordset2.CursorLocation = 2
			Recordset2.LockType = 1
			Recordset2.Open()
			Recordset2_numRows = 0
			If Recordset2.EOF then

			'Response.write "CRAP!  I GOT HERE TOO!<BR>"
			'Response.write "VEhicleID="&VehicleID&"<BR>"
			'Response.write "VehicleName="&VehicleName&"<BR>"
			'Response.write "VehicleType="&VehicleType&"<BR>"
			'Response.write "UserID="&UserID&"<BR>"

			Set RSEVENTS2 = Server.CreateObject("ADODB.Recordset")
				RSEVENTS2.Open "AvailableVehicles", Database, 2, 2
				RSEVENTS2.addnew	
				RSEVENTS2("VehicleID") = VehicleID
				RSEVENTS2("VehicleName") = VehicleName
                RSEVENTS2("VehicleType") = VehicleType
                RSEVENTS2("DriverID")=UserID	
				RSEVENTS2("LogInTime")=Now()		
				RSEVENTS2("AvailableStatus") = "c"
				RSEVENTS2.update
				RSEVENTS2.close			
			set RSEVENTS2 = nothing	



			End if
			Recordset2.Close()
			Set Recordset2 = Nothing
			else
			'Response.write "LoginCheck="&LoginCheck&"<BR>"
			IF LoginCheck<>"n" then
				'Response.write "WHY AM I HERE?!?<BR>"
				Response.redirect("driverlogin.asp")
			End if
End if


			'Response.write "VEhicleID="&VehicleID&"<BR>"
			'Response.write "VehicleName="&VehicleName&"<BR>"
			'Response.write "VehicleType="&VehicleType&"<BR>"
			'Response.write "UserID="&UserID&"<BR>"


%>
<link rel="stylesheet" type="text/css" href="../css/style.css">
