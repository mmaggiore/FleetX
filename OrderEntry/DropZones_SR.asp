<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title></title>
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../fleetexpress.inc" -->
<link rel="stylesheet" type="text/css" href="../css/Style.css">
<title>Valid Stockroom Destinations/Codes</title>
</head>
<body>
<table border="1">
<tr><td><b>Code</b></td><td><b>Name</b></td></tr>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionTimeout = 100
oConn.Provider = "MSDASQL"
oConn.Open DATABASE
	'l_cSQL = "select * FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id " &_
	'"WHERE (fcshipbt.sb_bt_id = '26') order by st_id"
    l_cSQL = "select * " &_
	"FROM PreExistingCompanies " &_
	"WHERE (isStockRoom = 'y') AND CompanyStatus='c' order by st_id"


	'Response.write "l_cSQL="&l_cSQL&"<BR>"
	SET oRs = oConn.Execute(l_cSql)
	Do while not oRs.EOF 
		St_ID=trim(oRs("st_ID"))
		txtPUCompany=trim(oRs("CompanyName"))
		txtPUContact=trim(oRs("ContactName"))
		txtPUPhone=trim(oRs("CompanyPhone"))
		fl_sf_addr1=trim(oRs("CompanyAddress"))
		fl_sf_addr2=trim(oRs("CompanySuite"))
		fl_sf_city=trim(oRs("CompanyCity"))
		fl_sf_state=trim(oRs("CompanyState"))
		txtPUZip=trim(oRs("CompanyZip"))
		%>
		<tr><td><%=St_ID%></td><td><%=txtPUCompany%></td></tr>
		<%
	oRs.Movenext
	Loop
Set oConn=Nothing
%>
<tr><td>&nbsp;</td></tr>
</table>




</body>
</html>
