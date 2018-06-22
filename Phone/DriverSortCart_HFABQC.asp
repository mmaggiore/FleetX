<%@ Language=VBScript %>
<!-- 10/08/2005 Span-Renuka included this screen for CCF 2023 -->
<!-- #include file="../DedicatedFleets/include/checkstring.inc" -->
<!-- #include file="../DedicatedFleets/include/custom.inc" -->
<!-- #include file="../DedicatedFleets/include/ifabsettings.inc" -->
<!-- #include file="driverinfo.inc" -->	
<HTML>
	<HEAD>
		<!--meta http-equiv="refresh" content="100"-->
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<script type="text/javascript">
		function formSubmit()
		{
		document.getElementById("thisForm").submit()
		}
		</script>
		<script type="text/javascript">
		    function formSubmit11() {
		        document.getElementById("thisForm11").submit()
		    }
		</script>				
		<%
		varFromLocations=" or fl_sf_id='72' "
		OtherBillToID=Request.Cookies("Phone")("sBT_ID")	
		fh_bt_id=Request.Cookies("Phone")("sBT_ID")
		If Request.Form("page") = "" Then
			intPage = 1	
			Else
			intPage = Request.Form("page")
		End If	
		If Request.Form("page2") = "" Then
			intPage2 = 1	
			Else
			intPage2 = Request.Form("page2")
		End If	
		ScannedLot=Request.Form("ScannedLot")
		rf_box=Request.Form("Rf_Box")		
		AcknowledgeIt=Request.Form("AcknowledgeIt")
		AliasCode=Request.Form("AliasCode")
		If AliasCode="" then
			AliasCode=Request.QueryString("AliasCode")
		End if
		If AliasCode>"" then Response.Cookies("Phone")("AliasCode")=AliasCode end if
		If aliasCode="" then aliasCode=Request.Cookies("Phone")("AliasCode") end if
		LocationCode=Request.Form("LocationCode")
		FakeSubmit=Request.Form("FakeSubmit")
		If FakeSubmit="" then
			FakeSubmit=Request.QueryString("FakeSubmit")
		End if
		If FakeSubmit>"" then
			Response.Cookies("Phone")("FakeSubmit")=FakeSubmit
		End if
		If FakeSubmit="" then
			FakeSubmit=Request.Cookies("Phone")("FakeSubmit")
		End if
		PageStatus=Request.Form("PageStatus")

		txtJobNumber=Request.Form("txtJobNumber")
		If FakeSubmit="fakesubmit" then
			Set Recordset1 = Server.CreateObject("ADODB.Recordset")
			Recordset1.ActiveConnection = DATABASE
			Recordset1.Source = "SELECT st_id, sb_bt_id FROM fcshipto INNER JOIN fcshipbt ON fcshipto.st_id = fcshipbt.sb_st_id WHERE (st_alias='"&AliasCode&"')"
			'response.write "Recordset1.Source="&Recordset1.Source&"<BR>"
			Recordset1.CursorType = 0
			Recordset1.CursorLocation = 2
			Recordset1.LockType = 1
			Recordset1.Open()
			Recordset1_numRows = 0
			If Recordset1.eof then
				ErrorMessage="That is not a valid location"
			End if			
			Do While NOT Recordset1.EOF 
			'''''''''''''CHANGED to LOOP so locations with the same alias code would work correctly''''''''''''
				SetArrivalTime="y"
				m=m+1
				LocationCode=Recordset1("st_id")
				BillToID=Trim(cStr(Recordset1("sb_bt_id")))
				varToLocations=varToLocations&" or fl_st_id='"&trim(LocationCode)&"' "
				varFromLocations=varFromLocations&" or fl_sf_id='"&trim(LocationCode)&"' "
				If OtherBillToID="80" then
					BillToID="80"
				End if
				Recordset1.Movenext
				Loop
					Response.Write "</font>"
			Recordset1.Close()
			Set Recordset1 = Nothing
			LengthvarToLocations=len(varToLocations)
			LengthvarFromLocations=len(varFromLocations)
			'Response.Write "varToLocations="&varToLocations&"<BR>"
			'Response.Write "varFromLocations="&varFromLocations&"<BR>"
			'Response.Write "LengthvarToLocations="&LengthvarToLocations&"<BR>"
			'Response.Write "LengthvarFromLocations="&LengthvarFromLocations&"<BR>"
			If m>0 then
			    varToLocations="("&Right(varToLocations, (int(LengthvarToLocations)-3))&")"	
			    varFromLocations="("&Right(varFromLocations, (int(LengthvarFromLocations)-3))&")"	
			End if	
			
			AliasCode=UCASE(ALIASCODE)
			LocationCode=Trim(UCASE(LOCATIONCODE))
			DisplayLocationCode=LocationCode
		If ErrorMessage="" then PageStatus="loggedin" End if
		End if
		If PageStatus>"" then 
			Response.Cookies("Phone")("PageStatus")=PageStatus 
		end if
		If PageStatus="" then 
			PageStatus=Request.Cookies("Phone")("PageStatus") 
		end if
		''''''''''''''''''''HERE'S WHERE I START CLOSING OUT THE SORTED WAFERS'''''''''''''''''''''''
		SQLStuff=Request.Form("SQLStuff")
		If trim(SQLStuff)="y" then
		        ''''''''''''''SEE IF THE LOT ACTUALLY EXISTS
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				SQL = "SELECT rf_ref, fl_sf_id, fl_st_id , fh_id "
				SQL = SQL&"FROM fcrefs INNER JOIN fcfgthd ON fcrefs.rf_fh_id = fcfgthd.fh_id INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id "
				SQL = SQL&"WHERE (fcfgthd.fh_status = 'UNS') AND (fcrefs.ref_status IS NULL) "
				SQL = SQL&"AND (fcrefs.rf_box = '"&rf_box&"') and (fcrefs.rf_ref = '"& ScannedLot &"')  AND (fclegs.fl_leg_status = 'f') "
				SQL = SQL&"ORDER BY fcrefs.rf_ref "			
				'Response.Write "SQL="&SQL&"<BR>"
				oRs.Open SQL, DATABASE, 1, 3
				IF NOT oRS.EOF THEN
				    Origination=trim(oRS("fl_sf_id"))
				    Destination=trim(oRS("fl_st_id"))
				    JobNumber=trim(oRS("fh_id"))
				    DisplayMessage="<font color='BLUE'>Destination: "& Destination &"</font>"
				    ''''''''UPDATES THE REF
				    Set oConn64 = Server.CreateObject("ADODB.Connection")
				    oConn64.ConnectionTimeout = 100
				    oConn64.Provider = "MSDASQL"
				    oConn64.Open DATABASE
				    l_cSQL64 = "UPDATE FCREFS SET ref_status = 's'" 
                    l_cSQL64 = l_cSQL64&" WHERE (rf_ref='"& scannedlot &"') AND (rf_box='"& rf_box &"') AND (ref_status is NULL)"
				    'Response.Write "****l_cSQL64="& l_cSQL64 &"<BR>"
				    oConn64.Execute(l_cSQL64)
				    oConn64.Close
				    Set oConn64=Nothing
				    
				    
	                Set Recordset1 = Server.CreateObject("ADODB.Recordset")
                    SQL777="SELECT rf_ref FROM FCREFS WHERE (rf_box='"& rf_box &"') AND (rf_fh_id='"& JobNumber &"') AND ((ref_status<>'s') OR (ref_status is NULL))"
                    'Response.Write "<br>SQL777="&SQL777&"***<BR>"
                    Recordset1.ActiveConnection = Database
                    Recordset1.Source = SQL777
                    Recordset1.CursorType = 0
                    Recordset1.CursorLocation = 2
                    Recordset1.LockType = 1
                    Recordset1.Open()
                    Recordset1_numRows = 0
	                if NOT Recordset1.EOF then
		                    whatever=Recordset1("rf_ref")
		                 Else
		                    '''''''''CHANGE THE JOB STATUS AND ROUTE IT''''''''
			        Set oConn = Server.CreateObject("ADODB.Connection")
			        oConn.ConnectionTimeout = 100
			        oConn.Provider = "MSDASQL"
			        oConn.Open DATABASE			                    
		            L_SQL_44="PHONE_CHANGE_STATUS '" & JobNumber & "', '53', 'ARV', '', '',  '"& UserID &"', '"& UnitID &"'" 
		           'Response.Write "L_SQL_44(2)="& L_SQL_44 &"<BR>"
		            oConn.Execute(L_SQL_44)		                    
 		            oConn.Close
		            Set oConn=Nothing  		                    
		                    'ErrorMessage="Incorrect second driver ID or password"
		            End if
		            Recordset1.Close()
		            Set Recordset1 = Nothing			    
				    
				    
				    				
				    '''''''''''''''''''''''
				    'Response.Write "GOT HERE!!!!<BR>"
				    'ZE=ZE+1
                    'Rf_ref=oRS("RF_REF")
                    '''''''''SHOULD THE ORDER STATUS CHANGE???''''''
                    
                    
                    Else
                    DisplayMessage="<font color='red'>Lot #"& ScannedLot &" should not be in this cart.<br>Please contact your supervisor.</font>"
				End if
				oRs.Close
				Set oRs=Nothing	
				
						
		        'Response.Write "Hola?<BR>"
		    
 		    
		End if
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		%>
	</HEAD>
	<%if pagestatus>"" then%>
		<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.thisForm11.scannedlot.focus()>
		<%else%>
		<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.thisForm.rf_box.focus()>
	<%end if%>	
		<%
		
		Select Case PageStatus
			Case "loggedin"
'-------------------STARTS THE DROP OFF				
				
				X=0
				Y=0
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.CursorLocation = 3
				oRs.CursorType = 3
				oRs.ActiveConnection = DATABASE	
				SQL = "SELECT rf_ref "
				SQL = SQL&"FROM fcrefs INNER JOIN fcfgthd ON fcrefs.rf_fh_id = fcfgthd.fh_id "
				SQL = SQL&"WHERE (fcfgthd.fh_status = 'UNS') AND (fcrefs.ref_status IS NULL) "
				SQL = SQL&"AND (fcrefs.rf_box = '"&rf_box&"') AND (fh_bt_id='36') "
				SQL = SQL&"ORDER BY fcrefs.rf_ref "			
				
				
				
				'SQL = "SELECT Fl_SF_ID, Fh_ID, fh_User5, Fl_ST_ID, fl_sf_comment, FH_Status, Fh_Priority FROM fcfgthd INNER JOIN fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id"
				'SQL = SQL&" WHERE (fh_ship_dt>'"&now()-30&"') AND ((FL_Leg_Status='c') OR (FL_Leg_Status is NULL)) AND (Fl_dr_ID='"&VehicleID&"') and ("
				'SQL = SQL&varToLocations
				'SQL = SQL&") AND (fh_status='ARV')"
				'SQL = SQL&" ORDER BY fh_priority, fh_id"
				
				oRs.Open SQL, DATABASE, 1, 3
				If trim(DisplayLocationCode)="55" then DisplayLocationCode="CPGP" end if
				If trim(DisplayLocationCode)="48" then DisplayLocationCode="KWEO" end if
				'Response.Write "****SQL="&SQL&"<BR>"
				%>
					<table width="300" cellpadding="0" cellspacing="0" border="0" bordercolor="green" align="left" ID="Table1">
						<tr><td align="center" colspan="3"><form method="post" action="default.asp" ID="Form7"><input type="submit" value="Return to Menu" ID="Submit1" NAME="Submit1"></form></td></tr>
						<tr>
							<td class="mainpagetextboldcenter" colspan="3" align="center">
								<font color="blue">Last update: <%=Time()%></font>
							</td>
						</tr>						
						<tr>
							<td align="center" class="purpleseparator" colspan="3"><b>WAFERS IN CART # <%=rf_box%></b></td>
						</tr>
						<tr><td>&nbsp;</td>
				<%
				If trim(DisplayMessage)>"" then
				    %>
    				<tr><td class="mainpagetextboldcenter" colspan="3" align="center"><%=DisplayMessage%></td></tr>
				    <tr><td>&nbsp;</td></tr>
				    <%
				End if
				If not oRs.EOF then
					%>
						<form method="post" ID="thisForm11" name="thisForm11">
						<tr>
							<td align="center" nowrap colspan="3"><input type="text" name="scannedlot" /></td>
						</tr>
						<input type="hidden" name="pagestatus" value="<%=pagestatus%>" ID="Hidden1">
						<input type="hidden" name="VehicleID" value="<%=VehicleID%>" ID="Hidden2">
						<input type="hidden" name="rf_box" value="<%=rf_box%>" ID="Hidden9">
						<input type="hidden" name="SQLStuff" value="y" />
						<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden3">
						<input type="hidden" name="truckstatus" value="<%=truckstatus%>" ID="Hidden4">						
						<input type="hidden" name="page" value="<%=intPage - 1%>" ID="Hidden5">
						<input type="hidden" name="AliasCode" value="<%=Trim(AliasCode)%>" ID="Hidden6">	
						<input type="hidden" name="Hub2" value="<%=Hub2%>" ID="Hidden7">
						<input type="hidden" name="Hub" value="<%=Hub%>" ID="Hidden8">	
					    <tr><td>&nbsp;</td></tr>
					    <tr><td colspan="2" align="center"><input size="8" maxlength="30" type="text" value="<%=Time()%>" name="bogus" ID="Text2" onFocus="formSubmit11()" readonly="readonly" class="InvisibleTextFieldPurple"></td></tr>							
						</form>						
						<tr><td>&nbsp;</td></tr>							
						<tr>
							<td align="center">&nbsp;</td>						
							<td align="left"><b>&nbsp;&nbsp;&nbsp;&nbsp;Unsorted Lots:</b>
						<%
						CloseTable="y"
						ELSE
						Response.Write "<tr><td colspan='4' align='center'>No lots to sort for this cart.</td></tr><tr><td>&nbsp;"
				End if

				Do while NOT oRS.EOF
				    ZE=ZE+1
                    Rf_ref=oRS("RF_REF")
                    If ZE>1 then
                        Response.Write ", "
                    End if
                    Response.Write trim(Rf_Ref)
				oRs.movenext
				LOOP
				oRs.Close
				Set oRs=Nothing						
					
				Response.Write "</font></td></tr>"

				
				
				

			Case else

			Set oConn=Nothing				
			%>
				<TABLE WIDTH="300" cellpadding="0" cellspacing="5" ID="Table2" align="left" border="0" bordercolor="red">
					<tr><td align="center" colspan="3"><form method="post" action="default.asp" ID="Form1"><input type="submit" value="Return to Menu" ID="Submit3" NAME="Submit1"></form></td></tr>			
				</table>
				<br clear="all">
			<FORM ACTION="DriverSortCart_HFABQC.asp" method="post" name="thisForm" ID="Form6">
				
				<TABLE WIDTH="300" cellpadding="0" cellspacing="5" ID="Table3" align="left" border="0" bordercolor="red">
					<TR> 
						<td> 
							<div class="purpleseparator"> 
								<table border="0" cellpadding="2" cellspacing="0" ID="Table4" width="100%" bordercolor="blue">
									<tr> 
										<td class="mainpagetextboldright" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
									</tr>
									<tr>
										<td class='mainpagetextboldcenter' colspan="2" nowrap align="center">SCAN in Cart ID</td>
									</tr>
									<tr>
										<td colspan='2' class='generalcontent' align="center">
											<input maxlength="20" name="rf_box" id="txtstation" type="text" size="15">
											<input maxlength='25' size='25' name='VehicleID' id='VehicleID' value='<%=VehicleID%>' type="hidden">
											<input type="hidden" name="FakeSubmit" value="fakesubmit" ID="Hidden16">
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan="2" align="center"><input size="8" maxlength="30" type="text" value="<%=Time()%>" name="bogus" ID="Text1" onFocus="formSubmit()" readonly="readonly" class="InvisibleTextFieldPurple"></td></tr>				
			
									<%if errormessage>"" then%>
										<tr>
											<td class='generalcontenthighlight'colspan='2' align="center"><font color="red"><br><b><%=ErrorMessage%></b><br><br></font></td>
										</tr>
									<%end if%>
									<tr> 
										<td class="subheader" colspan="2"><img src="../images/transpixel.gif" height="2"></td>
									</tr>
								</table>
							</div>
						</td>
						<!--Dummy section-->
					</TR>
					<tr><td align="center" colspan="4">&nbsp;</td></tr>					
				</TABLE>
			</FORM>
			
			<%
			End Select
			%>
			
	</BODY>
</HTML>
