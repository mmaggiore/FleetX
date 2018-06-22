<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>Log Out Page</title>
<!-- #include file="FleetX.inc" -->
</head>
<body>
<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"

	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.ConnectionTimeout = 100
	oConn.Provider = "MSDASQL"
	oConn.Open DATABASE
        'l_cSQL = "UPDATE AvailableVehicles SET availablestatus = 'x' WHERE DriverID = '" & UserID & "'"
        l_cSQL = "UPDATE AvailableVehicles SET logouttime=getdate(), availablestatus = 'x' WHERE DriverID = '" & UserID & "' and AvailableStatus='c'"
	oConn.Execute(l_cSQL)
	oConn.close
	Set oConn=Nothing
    VehicleName=""

Response.Cookies("Phone")("DriverEmail")=""
Response.Cookies("Phone")("DriverUserID")=""
Response.Cookies("Phone")("DriverFirstName")=""
Response.Cookies("Phone")("DriverLastName")=""
Response.Cookies("Phone")("Rights")=""
Response.Cookies("Phone")("VehicleID")=""
Response.redirect("DriverLogin.asp?x=1")
'Response.redirect("../phone.asp")
 %>

</body>
</html>
