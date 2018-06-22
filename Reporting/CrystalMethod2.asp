<%@ Language=VBScript CodePage=65001  %>

<%
'==========================================================================
' CRYSTAL ENTERPRISE REPORT APPLICATION SERVER (CE EMBEDDED) 10
' Purpose:  Demonstrate how to pass parameters to subreports
'==========================================================================


' DC making this a "Multi-Call" form with a hidden object
IF Request.Form("hdnHaveParms") = "YES" Then
	' We're coming here the second time and, therefore, have parameter values

	' This is the Radio Button Selection
	l_cSel = Request.Form("STSEL")
	
	' This is the DATE the user entered
	l_cDate = Request.Form("txtDate")                                                                  
	l_cEndDate = Request.Form("strEnd")
	
	'Thiss line creates a string variable called reportname that we will use to pass
	'the Crystal Report filename (.rpt file) to the OpenReport method.
	reportname = "CycleTime_New_Wafer.rpt"



'============================================================================
' CREATE THE REPORT CLIENT DOCUMENT OBJECT AND OPEN THE REPORT
'============================================================================

' Use the Object Factory object to create other RAS objects (useful for versioning changes)
Set objFactory = CreateObject("CrystalReports.ObjectFactory")
	
' This "While/Wend" loop is used to determine the physical path (eg: C:\) to the 
' Crystal Report .rpt by translating the URL virtual path (eg: http://Domain/Dir)   
Dim path, iLen

path = Request.ServerVariables("PATH_TRANSLATED") 

While (Right(path, 1) <> "\" And Len(path) <> 0)                      
	iLen = Len(path) - 1                                                  
	path = Left(path, iLen)                                               
Wend       
                                                                                                                                  
' Create a new ReportClientDocument object
Set Session("oClientDoc") = objFactory.CreateObject("CrystalClientDoc.ReportClientDocument")

' Specify the RAS Server (computer name or IP address) to use (If SDK and RAS Service are running on seperate machines)
' 192.168.111.2
Session("oClientDoc").ReportAppServer = Session("ReportAppServerVariable")

' Open the report object to initialize the ReportClientDocument
Session("oClientDoc").Open path & reportName  


'==================================================================
' WORKING WITH DISCRETE PARAMETERS
'==================================================================

'This subroutine is called each time you have to pass a value to a Crystal Report parameter
Public Sub SetParamValues (ReportName, ParamName, ParamValue)

	'Creates a values collection to contain a list of parameters (discrete parameters in this case)
	Set Values = CreateObject("crystalReports.Values")
	
	'All the parameters in this example will be discrete parameters.  Therefore we create 
	'this object to hold the value of the parameter
	Set DiscreteValue = CreateObject ("CrystalReports.ParameterFieldDiscreteValue")

	'Sets the value of the discrete parameter
	DiscreteValue.Value = ParamValue
	
	'Adds the discrete parameter to the values collection
	Values.Add DiscreteValue 

	'Sets the current values for the parameter
	Session("oClientDoc").DataDefController.ParameterFieldController.SetCurrentValues ReportName, ParamName, Values

End Sub

	
			
		

	    
SetParamValues "", "FromDate", CDate(l_cDate)	    
	' SEcond param
 SetParamValues "", "Billto_ID",  CStr(Session("sUid"))
SetParamValues "", "ToDate", CDate(l_cEndDate)		



'============================================================================
' CHOOSING THE REPORT VIEWER
'============================================================================
'
' There are four Report Viewers:
' 1.  Crystal Reports Interactive Viewer (CrystalReportsInteractiveViewer.asp)
' 2.  Crystal Reports Viewer (CrystalReportsViewer.asp)
' 3.  Crystal Reports Parts Viewer (CrystalReportsPartsViewer.asp)
' 4.  Legacy ActiveX Viewer (ActiveXViewer.asp)
'
' Note that to use this these viewers you must have the appropriate .asp file in the 
' same virtual directory as the main ASP page. Choose from one of the four viewers below,
' simply uncomment the one you want to use:

'  *** Crystal Reports Interactive Viewer ***
Response.Redirect "CrystalReportsInteractiveViewer.asp"

'  *** Crystal Reports Viewer ***
'Response.Redirect "CrystalReportsViewer.asp"

'    IMPORTANT NOTE - To use the report parts viewer successfully you are required to
'    choose and name three objects in the report to Node1, Node2 and Node3.
'    You can access an objects name by using the Format Editor dialog box.
'    For more information on the Format Editor Dialog Box and setting objects
'    names, please refer to the Help Contents (Help Menu->Crystal Reports Help)
'    or by pressing F1

'  *** Crystal Reports Parts Viewer ***
'Response.Redirect "CrystalReportsPartsViewer.asp"

' 	 IMPORTANT NOTE - The ActiveXviewer does NOT have the ability to prompt for parameters,
' 	 selection formulas, or login information.  If your report requires this information
'    to run, look at the the appropriate code samples to pass this information before
' 	 using this viewer.

'  *** Legacy ActiveX Viewer ***
'Response.Redirect "ActiveXViewer.asp"

'=============================================================================
ELSE
	' This is the first time and we need to prompt for parameters
	%>
	<html><head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<TITLE><% Response.Write(D_TITLEBAR) %></TITLE>
	
		<!-- 08/17/2004 Span-Renuka - Added this code for displaying calender control accordingly-->
	<% currpath=Request.ServerVariables("PATH_INFO")%>
	<% IF currpath="/Reporting/CrystalMethod2.asp" THEN %>
		<SCRIPT LANGUAGE="JavaScript" SRC="../v9web/scripts/date-picker.js"></SCRIPT>
	<%ELSE %>
		<SCRIPT LANGUAGE="JavaScript" SRC="../scripts/date-picker.js"></SCRIPT>
	<%END IF %>
		<!-- 08/17/2004 Span-Renuka commented this code -->	
	</head>
	<%
	IF D_CBODYELEMENTS <> "" THEN
		' They have specified some custom body elements
		Response.Write(" " & D_CBODYELEMENTS & " ")
	END IF
	%>
	<CENTER>
	<%
	' Response.Write("<IMG SRC='" & Session("sLogo") & "'>")

	' Optionally show the Bold Heading of the purpose of this page
	Response.Write("<H2><FONT COLOR=#660099>Jobs Report</FONT></H2>")

	IF SHOW_BTNAME THEN
		Response.Write("<H2>" & D_CINETFROM & " " & Session("txt_cm_desc"))
		IF SHOW_BTID THEN
			Response.Write(" (Customer ID:" & Session("suid") & ")")
		END IF
		Response.Write("</H2><BR>")
	END IF
	
	'  Want to have:
	'	(*) Yesterday's Jobs
	'	(*) Today's Jobs
	'	(*) Specific Date:  [mm/dd/yy]

	%>
	<!-- #include file="../include/settings.inc" -->

	<form name="GetJobParms" action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="POST">
	<table border="4" cellpadding="2" cellspacing="0">
	<tr>
	<td>
	<table border="0" bordercolor="red" width="350">
	<!--tr><td><img src="../images/transpixel.gif" width="1" height="10"></td></tr-->
    <TR><td class="generalcontent" align="center">Please choose your report selection criteria</td></TR>
    
	<form name="GetJobParms" action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="POST" ID="Form1">

<TR> 
    <td align="center"> 
          <table border=0 cellpadding="2" cellspacing="0" ID="Table1">
            <tr> 
              <td class="subheader" colspan="4"><img src="../images/transpixel.gif" height="2"></td>
            </tr>
            <tr> 
              <td class="generalcontent" width="97">From</td>
              <td width="696" class="generalcontent"> 
                <input type='text' size='12' name='txtDate' value='<%=Date()-1%>' maxlength="12" ID="Text1">
                &nbsp;<a href="javascript:show_calendar('GetJobParms.txtDate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;"><img src="../images/cal.gif" width="16" height="15" border="0" name="calendar" alt="Calendar" title="Calendar" align="ABSMIDDLE"></a>
                </td>
                
                
                
              <td class="generalcontent" width="97">To</td>
              <td width="696" class="generalcontent"> 
                <input type='text' size='12' name='strEnd' value='<%=Date()-1%>' maxlength="12" ID="Text2">
                &nbsp;<a href="javascript:show_calendar('GetJobParms.strEnd');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;"><img src="../images/cal.gif" width="16" height="15" border="0" name="calendar" alt="Calendar" title="Calendar" align="ABSMIDDLE"></a>
                </td>              
                
                
                
            </tr>
            <tr> 
              <td class="subheader" colspan="4"><img src="../images/transpixel.gif" height="2"></td>
            </tr>
            </table>
            </td></tr>
            
            </table>
      </td>
    </TR>
	</table>
	<BR>
	<input type="hidden" name="hdnHaveParms" value="YES">
	<input type="submit" name="btnSubmit" value="View Report">
	<input type="reset" name="btnReset" value="Reset">
	</form>
	</CENTER>
	<%
END IF
	%>