		
		<%
		'''''''''ID COOKIE INFO'''''''''''
		ComputerID=Request.Cookies("LegalComputer")("ComputerID")
        'Response.write "computerid="&computerid&"<BR>"


PhoneUser=Request.ServerVariables("HTTP_UA_OS") 

'Response.write "PhoneUser="&PhoneUser&"<BR><br>"
 
'For Each name In Request.ServerVariables  
'response.write(name&": ") 
'response.write(Request.ServerVariables(name)&"<br><br>")
'Next





		'If trim(ComputerID)="" and Trim(PhoneUser)<>"Windows CE (Pocket PC) - Version 5.1" then
		'	RequireCookie="y"
		'	Response.write "<font color='red'>12/14/11-If you are using a LogistiCorp<br>phone to access this page, and you are<BR>seeing this message,<BR>then please call Mark Maggiore<BR>at 214-956-0400 xt. 212, ASAP.</font>"
		'End if


		DriverEmail=Request.Cookies("FleetXPhone")("DriverEmail")
		UserID=Request.Cookies("FleetXPhone")("DriverUserID")
		FirstName=Request.Cookies("FleetXPhone")("DriverFirstName")
		LastName=Request.Cookies("FleetXPhone")("DriverLastName")
		Rights=Request.Cookies("FleetXPhone")("Rights")
		VehicleID=Request.Cookies("FleetXPhone")("VehicleID")
		'Response.write "VehicleID="&VehicleID&"<BR>"
		'Response.write "VehicleID="&VehicleID&"<BR>"
		VehicleName=Request.Cookies("FleetXPhone")("VehicleName")
		VehicleType=Request.Cookies("FleetXPhone")("VehicleType")
		UnitID=Request.Cookies("FleetXPhone")("UnitID")
		VehicleID=Request.Cookies("FleetXPhone")("VehicleID")
		'response.write "hello?<BR>"
		'Response.write "*********VehicleID="&VehicleID&"<BR>"
		'Response.write "*********VehicleName="&VehicleName&"<BR>"
		'Response.write "XZXZXZXZUnitID="&UnitID&"<BR>"
			'Response.Cookies("FleetXPhone")("UnitID")=VehicleID
			'Response.Cookies("FleetXPhone")("VehicleID")=PenchantVehicleID
			'Response.Cookies("FleetXPhone")("VehicleName")=VehicleName		

url = Request.ServerVariables ("URL")
arrayX = split(url,"/")
PageName=lcase(ArrayX(ubound(arrayX)))
'Response.write "LINE 47 - "&PageName



		
		
		UnitID=Request.Cookies("FleetXPhone")("UnitID")
		If (trim(UserID)="" or trim(VehicleID)="") and PageName<>"drivervehicle.asp"  then
			'Response.write "DriverInfo.inc line 48<BR>"
			'Response.Redirect("DriverLogin.asp")
		End if	
		TempJobNumber=Request.Cookies("FleetXPhone")("TempJobNumber")	
		Response.Cookies("FleetXPhone")("TempJobNumber")=""
		'Response.write "UnitID="&UnitID&"<BR>"
		'Response.write "UserID="&UserID&"<BR>"
		'Response.write "VehicleID="&VehicleID&"<BR>"
		If UnitID="303551" or UnitID="303552" or UnitID="303553" or UnitID="303554" then
			DisplayP0="y"
		End if
		showtop=session("showtop")
		'Response.write "showtop="&Showtop&"<BR>"
		SomeVariable=Request.ServerVariables ("HTTP_USER_AGENT")
        'Response.Write "SomeVariable="& SomeVariable &"<BR>"
        If left(SomeVariable,15)="Motorola_ES405B" then
            'Response.Write "YES<BR>"
            %>
            <META NAME="MobileOptimized" CONTENT="0">
            <%
            Else
            'Response.Write "NO<BR>"
        End if
		%>
		<meta name="viewport" content="width=300">
<%If showtop="y" then %>
<!-- #include file="../../dedicatedfleets/nav/ifabnavbar.inc" -->
<%end if 
varBTID=Request.Cookies("FleetXPhone")("sBT_ID")	
'Response.write "varBTID="&varBTID&"<BR>"

If trim(userID)="1" then
Set objWSHNetwork = Server.CreateObject("WScript.Network") 
Response.Write mid(objWSHNetwork.ComputerName,8,1)
End if
'''''''TI DATABASE INFO FOR TI SHERMAN JOBS
DATABASE3="DATABASE=fc_mdt;DSN=SQLConnect;UID=sa;Password=cadre;"
%>		