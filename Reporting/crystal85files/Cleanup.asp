<HTML>
<HEAD>
<TITLE>Session Cleanup</Title>
</HEAD>
<BODY Onload="CallClose();">
<%
'The javascript code is used to close the newly opened browser window once the page has
'run

Session.Contents.Remove("oPageEngine")
Session.Contents.Remove("oRpt")
Session.Contents.Remove("oApp")

'These last few lines remove the Application, Report and PageEngine session variables
'from the session variables collection to release the RDC license.
%>
<SCRIPT LANGUAGE="Javascript">
function CallClose()
{
window.close();
}
</SCRIPT>
</BODY>
</HTML>

