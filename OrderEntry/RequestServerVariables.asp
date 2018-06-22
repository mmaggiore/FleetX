<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>Untitled Page</title>
</head>
<body>
<%
for each x in Request.ServerVariables
  Y=Request.ServerVariables (x)
  Response.write x
  Response.write Y&"<BR><BR>"
next
%>
</body>
</html>
