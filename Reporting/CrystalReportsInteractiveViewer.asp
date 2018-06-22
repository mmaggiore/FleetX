<%@ Language=VBScript %>

<%
'=============================================================================
'INSTANTIATE THE VIEWER AND DISPLAY THE REPORT THROUGH THE INTERACTIVE VIEWER
'=============================================================================
Response.ExpiresAbsolute = Now() - 1
	
' Create the Crystal Reports Interactive Viewer
Dim viewer
Set viewer = CreateObject("CrystalReports.CrystalReportInteractiveViewer")  
viewer.Name = "Crystal Reports Interactive Viewer"
viewer.IsOwnForm = True	  
viewer.IsOwnPage = True
viewer.PageTitle = "Crystal Reports"
viewer.IsDisplayGroupTree = False
viewer.HasToggleGroupTreeButton = True

' NEW - Set the printmode to ActiveX printing
' Acceptable values: 0 (PDF printing), 1 (ActiveX printing)
viewer.PrintMode = 1

' IMPORTANT NOTE:
' For a complete list of properties of the Interactive Viewer look in the RAS "COM Viewer SDK"
' help file found through Start | Programs | Crystal Enterprise 10 | Crystal Enterprise Developer Documentation

' Set the source for the viewer to the ReportClientDocuments report source
viewer.ReportSource = Session("oClientDoc").ReportSource

' Optional: Add Search Control functionality to the Viewer
Dim BooleanSearchControl
Set BooleanSearchControl = CreateObject("CrystalReports.BooleanSearchControl")
BooleanSearchControl.ReportDocument = Session("oClientDoc")
viewer.BooleanSearchControl = BooleanSearchControl

' Process the http request to view the report
viewer.ProcessHttpRequest Request, Response, Session
%>