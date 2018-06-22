<%@ Language=VBScript %>
<!-- #include file="FleetX.inc" -->
<!-- #include file="driverinfo.inc" -->

<%
DriverID = Request.Cookies("FleetXPhone")("DriverUserID")
jobNum = request.querystring("j")
accCharge = request.querystring("d")
accType = request.querystring("c")
accID = request.querystring("a")
billTo = request.querystring("b")
jtype = request.querystring("t")
LocationCode = request.querystring("l")

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 100
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
 
SQL = "INSERT into ChargedAccessorials values('" & jobNum & "', '" & billTo & "'," & accType & "," & accID & "," & accCharge & "," & DriverID & ",'" & Now() & "','" & jtype & "','" & LocationCode & "')"
response.write "sql=" & SQL & "<Br>"

SET oRsN1 = oConn.Execute(SQL)

set oRsN1 = Nothing
Set oConn=Nothing

response.redirect "JobException.asp?j=" & jobNum & "&s=" & jtype & "&l=" & LocationCode
%>
