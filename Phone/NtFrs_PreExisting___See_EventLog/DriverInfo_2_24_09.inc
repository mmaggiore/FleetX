		<%
		DriverEmail=Request.Cookies("Phone")("DriverEmail")
		UserID=Request.Cookies("Phone")("DriverUserID")
		FirstName=Request.Cookies("Phone")("DriverFirstName")
		LastName=Request.Cookies("Phone")("DriverLastName")
		Rights=Request.Cookies("Phone")("Rights")
		VehicleID=Request.Cookies("Phone")("VehicleID")
		'Response.write "VehicleID="&VehicleID&"<BR>"
		'Response.write "VehicleID="&VehicleID&"<BR>"
		VehicleName=Request.Cookies("Phone")("VehicleName")
		UnitID=Request.Cookies("Phone")("UnitID")
		VehicleID=Request.Cookies("Phone")("VehicleID")
		'response.write "hello?<BR>"
		'Response.write "*********VehicleID="&VehicleID&"<BR>"
		'Response.write "*********VehicleName="&VehicleName&"<BR>"
		'Response.write "XZXZXZXZUnitID="&UnitID&"<BR>"
			'Response.Cookies("Phone")("UnitID")=VehicleID
			'Response.Cookies("Phone")("VehicleID")=PenchantVehicleID
			'Response.Cookies("Phone")("VehicleName")=VehicleName		
		
		
		UnitID=Request.Cookies("Phone")("UnitID")
		If UserID="" then
			Response.Redirect("DriverLogin.asp")
		End if	
		TempJobNumber=Request.Cookies("Phone")("TempJobNumber")	
		Response.Cookies("Phone")("TempJobNumber")=""
		'Response.write "UnitID="&UnitID&"<BR>"
		'Response.write "UserID="&UserID&"<BR>"
		'Response.write "VehicleID="&VehicleID&"<BR>"
		If UnitID="303551" or UnitID="303552" or UnitID="303553" or UnitID="303554" then
			DisplayP0="y"
		End if
		%>