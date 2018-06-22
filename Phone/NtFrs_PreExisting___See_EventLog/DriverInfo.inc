		<%
		DriverEmail=Request.Cookies("FleetXPhone")("DriverEmail")
		UserID=Request.Cookies("FleetXPhone")("DriverUserID")
		FirstName=Request.Cookies("FleetXPhone")("DriverFirstName")
		LastName=Request.Cookies("FleetXPhone")("DriverLastName")
		Rights=Request.Cookies("FleetXPhone")("Rights")
		VehicleID=trim(Request.Cookies("FleetXPhone")("VehicleID"))
		'Response.write "VehicleID="&VehicleID&"<BR>"
		'Response.write "VehicleID="&VehicleID&"<BR>"
		VehicleName=Request.Cookies("FleetXPhone")("VehicleName")
		UnitID=Request.Cookies("FleetXPhone")("UnitID")
		VehicleID=Request.Cookies("FleetXPhone")("VehicleID")
		'response.write "hello?<BR>"
		'Response.write "*********VehicleID="&VehicleID&"<BR>"
		'Response.write "*********VehicleName="&VehicleName&"<BR>"
		'Response.write "XZXZXZXZUnitID="&UnitID&"<BR>"
			'Response.Cookies("FleetXPhone")("UnitID")=VehicleID
			'Response.Cookies("FleetXPhone")("VehicleID")=PenchantVehicleID
			'Response.Cookies("FleetXPhone")("VehicleName")=VehicleName		
		
		
		UnitID=Request.Cookies("FleetXPhone")("UnitID")
		If UserID="" then
			Response.Redirect("DriverLogin.asp")
		End if	
		TempJobNumber=Request.Cookies("FleetXPhone")("TempJobNumber")	
		Response.Cookies("FleetXPhone")("TempJobNumber")=""
		'Response.write "UnitID="&UnitID&"<BR>"
		'Response.write "UserID="&UserID&"<BR>"
		'Response.write "VehicleID="&VehicleID&"<BR>"
		If UnitID="303551" or UnitID="303552" or UnitID="303553" or UnitID="303554" then
			DisplayP0="y"
		End if
		'response.write "hello?"
		%>