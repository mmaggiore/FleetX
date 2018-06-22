<html>
<head>

<!-- #include file="../fleetexpress.inc" -->
<%

    thisUser = Request.cookies("FleetXCookie")("UserID")
    CurrentDateTime=Now()

%>
</head>
<body>
   
<%
                Set oConnA = Server.CreateObject("ADODB.Connection")
                oConnA.ConnectionTimeout = 100
                oConnA.Provider = "MSDASQL"
                oConnA.Open DATABASE
                iuSQL = "Select TOP 1 * FROM AutoDispatchStatus ORDER BY adid DESC"
                'response.write "1702 sql=" & iuSQL & "<br>"
                SET oRsa2 = oConnA.Execute(iuSql)
                if oRsa2.eof then
                  ' first time, turn autodispatch ON
                  aSQL = "INSERT INTO AutoDispatchStatus (useron, dateon, status) VALUES(" & thisUser & ",'" & CurrentDateTime & "','c')"
                  SET oRsad = oConnA.Execute(aSql)
                  response.redirect "../home.asp"
                else
                  adstatus = trim(oRsa2("status"))
                  adtimeon = trim(oRsa2("dateon"))
                  adtimeoff = trim(oRsa2("dateoff"))
                  response.write "status = " & adstatus & "<br>"
                  if adstatus = "c" then
                    response.write "disable autodispatch -"
                    aSQL = "UPDATE AutoDispatchStatus SET status ='x', useroff =" & thisUser & ", dateoff='" & CurrentDateTime & "' WHERE adid=" & trim(oRsa2("adid"))
                    response.write aSQL & "<br>"
                    SET oRsad = oConnA.Execute(aSql)
                    response.redirect "../home.asp"
                  else
                    response.write "enable autodispatch ="
                    aSQL = "INSERT INTO AutoDispatchStatus (useron, dateon, status) VALUES(" & thisUser & ",'" & CurrentDateTime & "','c')"
                    response.write "aSQL=" & aSQL & "<Br>"
                    SET oRsad = oConnA.Execute(aSql)
                    response.redirect "../home.asp"
                  end if
                end if
                oRsa2.close
                Set oConnA=Nothing

%>
</body>
</html>

