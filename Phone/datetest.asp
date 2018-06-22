<%
indate = Now()
thisDay = datepart("D",DateAdd("D", 1, indate)) 
'response.write "thisday=" & thisDay
Response.Write "Today = " & FormatDateTime(indate, 2) & "<br>"
Response.write "Yesterday=" & FormatDateTime(DateAdd("D",-1,indate),2) & "<br>"
Response.write "Tomorrow=" & FormatDateTime(DateAdd("D",1,indate),2) & "<br>"
%>         