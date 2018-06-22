<%@ Language=VBScript %>
<!-- 10/08/2005 Span-Renuka included this screen for CCF 2023 -->
<!-- #include file="../v9web/include/checkstring.inc" -->
<!-- #include file="../v9web/include/custom.inc" -->
<!-- #include file="../v9web/include/ifabsettings.inc" -->
<!-- #include file="driverinfo.inc" -->	
<HTML>
	<HEAD>
		<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<link rel="stylesheet" type="text/css" href="mainStyleSheet.css">
		<SCRIPT type="text/javascript" language="JavaScript">
		function generateRow()
		{ 
			 //Read the number of Lotnumber text boxes present.
			 var numLinesAdded =document.thisForm.lotnumbertext.value ;
			 //Increment by 1
			 numLinesAdded++;
			 //Build the text box
			 //12/12/2005 Span-Renuka modified the code(Changed maxlength from 8 to 14) for allowing up to 13 characters.
			 var cellCnts1 = "<INPUT TYPE='text' maxlength='14' size=20 name='txtLotId" + numLinesAdded + "' onchange='validatedynamiclot(txtLotId" + numLinesAdded + ")'></td>";
			 numLinesAdded++;
			 var cellCnts2 = "<INPUT TYPE='text' maxlength='14' size=20 name='txtLotId" + numLinesAdded + "' onchange='validatedynamiclot(txtLotId" + numLinesAdded + ")'></td>";
			 numLinesAdded++;
			 var cellCnts3 = "<INPUT TYPE='text' maxlength='14' size=20 name='txtLotId" + numLinesAdded + "' onchange='validatedynamiclot(txtLotId" + numLinesAdded + ")'></td>";
			 numLinesAdded++;
			 var cellCnts4 = "<INPUT TYPE='text' maxlength='14' size=20 name='txtLotId" + numLinesAdded + "' onchange='validatedynamiclot(txtLotId" + numLinesAdded + ")'></td>";
			 numLinesAdded++;
			 var cellCnts5 = "<INPUT TYPE='text' maxlength='14' size=20 name='txtLotId" + numLinesAdded + "' onchange='validatedynamiclot(txtLotId" + numLinesAdded + ")'></td>";
			 var cellCnts6 = "&nbsp;" ;
			 //Read the number of rows present in the table
			 noRows = document.all.lottable.rows.length;
			 //Insert a new Row
			 tr = document.all.lottable.insertRow();
			 //Insert a cell         
			 td = document.all.lottable.rows(noRows).insertCell();
				document.all.lottable.rows(noRows).cells(0).innerHTML = cellCnts1;
			 td = document.all.lottable.rows(noRows).insertCell();
				document.all.lottable.rows(noRows).cells(1).innerHTML = cellCnts2;
			 td = document.all.lottable.rows(noRows).insertCell();
				document.all.lottable.rows(noRows).cells(2).innerHTML = cellCnts3;
			 td = document.all.lottable.rows(noRows).insertCell();
				document.all.lottable.rows(noRows).cells(3).innerHTML = cellCnts4;
			 td = document.all.lottable.rows(noRows).insertCell();
				document.all.lottable.rows(noRows).cells(4).innerHTML = cellCnts5;
			td = document.all.lottable.rows(noRows).insertCell();
				document.all.lottable.rows(noRows).cells(5).innerHTML = cellCnts6;
			//Assign the number of text boxes present.
			document.thisForm.lotnumbertext.value = numLinesAdded;
		}
		// 09/27/2005 Span-Renuka added the function to validate Lot numbers text boxes (IFAB CCF- 2023) 
		function validatelot(objectname)
		{
			var val = trimtext(document.getElementById(objectname).value);
			if (val != "" )
			{	
				//12/12/2005 Span-Renuka modified the code(Changed maxlength from 8 to 14) for allowing up to 13 characters.
				//12/13/2005 AMB modified the code to allow 5-13 characters
				//if (val.length != 7)
				if (val.length < 5 || val.length > 13)
				{
					//12/12/2005 Span-Renuka modified the code(Changed maxlength from 8 to 14) for allowing up to 13 characters.
					//alert("Invalid lot number. The lot number should have exactly 7 characters.");
					alert("Invalid lot number. The lot number should have 5-13 characters.");
					document.getElementById(objectname).value ="";
					document.getElementById(objectname).focus();
				}
				else
				{
					var lotsscanned =document.thisForm.lotnumbersscanned.value ;
					 lotsscanned++;
					document.thisForm.lotnumbersscanned.value = lotsscanned;
					document.thisForm.txtPieces.value = document.thisForm.lotnumbersscanned.value;
			    }
			} 
		}
		//09/23/2005 Span-Renuka added the function to validate dynamically generated Lot numbers text boxes (IFAB CCF- 2023) 
		function validatedynamiclot(objectname)
		{
			var val = trimtext(document.getElementById(objectname.name).value);
			if (val != "" )
			{
				//12/12/2005 Span-Renuka modified the code(Changed maxlength from 8 to 14) for allowing up to 13 characters.
				//if (val.length != 7)
				if (val.length < 5 || val.length > 13)
				{
					//12/12/2005 Span-Renuka modified the code(Changed maxlength from 8 to 14) for allowing up to 13 characters.
					//alert("Invalid lot number. The lot number should have exactly 7 characters.");
					alert("Invalid lot number. The lot number should have 5-13 characters.");
					document.getElementById(objectname.name).value ="";
					document.getElementById(objectname.name).focus();
				}
				else
				{
					var lotsscanned =document.thisForm.lotnumbersscanned.value ;
					lotsscanned++;
					document.thisForm.lotnumbersscanned.value = lotsscanned;
					document.thisForm.txtPieces.value = document.thisForm.lotnumbersscanned.value;
			    }
			} 
		}
		function trimtext(textval)
		{
			x = textval;
			while (x.substring(0,1) == ' ') x = x.substring(1);
			while (x.substring(x.length-1,x.length) == ' ') x = x.substring(0,x.length-1);
			return(x);
		}
		function updjob()
		{
			if (document.thisForm.txtPieces.value == document.thisForm.lotnumberstobescanned.value)
			{
				document.thisForm.savemode.value = "SAVE";
				document.thisForm.submit();
			}
			else
			{
				alert("Number of Lots were not matching")
			}
		}
	</SCRIPT>
	</HEAD>
	<%
	DriverID=Request.Form("DriverID")
	LocationCode=Request.Form("LocationCode")
	txtjobnumber=Request.Form("txtjobnumber")
	
			'Response.write "LocationCode="&LocationCode&"<BR>"
			'Response.write "DriverID="&DriverID&"<BR>"	
			'Response.write "******txtjobnumber="&txtjobnumber&"<BR>"	
	
	IF Request.Form("txtjobnumber") <> "" THEN 
		'Connection
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.ConnectionTimeout = 100
		oConn.Provider = "MSDASQL"
		oConn.Open DATABASE
		
		l_jobnumber = Trim(Request.Form("txtjobnumber"))
		l_stationid = Trim(Request.Form("txtstation"))
		'l_cSQLjob = "select * from fclegs where fl_fh_id='" & l_jobnumber & "'"
		l_cSQLjob = "SELECT fh_status, fl_sf_id, fl_sf_name, fl_st_id, fl_st_name, " & _
					"fl_numboxes, fh_statcode,fh_type FROM fclegs " & _
					"JOIN fcfgthd ON fl_fh_id = fh_id " & _
					"LEFT OUTER JOIN fcshipto csf ON fl_sf_id = csf.st_id " & _
					"WHERE fl_fh_id = '" & l_jobnumber & "' and fl_sf_id ='" & l_stationid & "'"				
		Set oRs = oConn.Execute(l_cSQLjob)
		'Response.write "THIS WOULD DETERMINE A PICKUP!!!!="&(l_cSQLjob)
		'Response.write "<br>GOT HERE #1<br>"
		'Response.End 
		IF not oRs.EOF then
			l_continue = TRUE
			l_pickup =TRUE
			l_fh_status = Trim(oRs.Fields("fh_status"))
			l_fl_sf_id = Trim(oRs.Fields("fl_sf_id"))
			l_fl_sf_name = Trim(oRs.Fields("fl_sf_name"))
			l_fl_st_id = Trim(oRs.Fields("fl_st_id"))
			l_fl_st_name = Trim(oRs.Fields("fl_st_name"))
			l_fl_numboxes = Trim(oRs.Fields("fl_numboxes"))
			l_fh_statcode = Trim(oRs.Fields("fh_statcode"))
		Else
			l_continue = FALSE
		END IF
		IF l_continue = FALSE THEN
			l_cSQLjob = "SELECT fh_status, fl_sf_id, fl_sf_name, fl_st_id, fl_st_name, " & _
						"fl_numboxes, fh_statcode,fh_type FROM fclegs " & _
						"JOIN fcfgthd ON fl_fh_id = fh_id " & _
						"LEFT OUTER JOIN fcshipto csf ON fl_sf_id = csf.st_id " & _
						"WHERE fl_fh_id = '" & l_jobnumber & "' and fl_st_id ='" & l_stationid & "'"				
			Set oRs = oConn.Execute(l_cSQLjob)
			
			'Response.write "l_cSQLjob="&l_cSQLjob
			
			IF not oRs.EOF then
				'Response.write "<br>GOT HERE #2<br>"
				l_continue = TRUE
				l_dropoff =TRUE
				l_fh_status = Trim(oRs.Fields("fh_status"))
				l_fl_sf_id = Trim(oRs.Fields("fl_sf_id"))
				l_fl_sf_name = Trim(oRs.Fields("fl_sf_name"))
				l_fl_st_id = Trim(oRs.Fields("fl_st_id"))
				l_fl_st_name = Trim(oRs.Fields("fl_st_name"))
				l_fl_numboxes = Trim(oRs.Fields("fl_numboxes"))
				l_fh_statcode = Trim(oRs.Fields("fh_statcode"))
			Else
				l_continue = FALSE
			END IF
		END IF	
	ELSE
		l_continue = FALSE
	END IF
	IF Request.Form("savemode") ="SAVE" THEN
		l_jobnumbersscanned = Request.Form("txtjobnumber")
		l_cPcs = Request.Form("txtPieces")
		'10/19/2005 Span-Renuka added the code for sending driver Id to Procedure.
		l_cDrid	= trim(Request.Form("txtcaller"))
		'10/17/2005 Span-Renuka added the code for checking the Lotnumbers.
		'Cross verifying the Lot numbers
		l_clotnumbertext = Request.Form("lotnumbertext")
		errorlot =""
		Previousjob =""
		For numblot = 1 to l_cPcs
			IF Trim(Request.Form("txtLotId"&numblot&"")) <> "" Then
				'Read each Lot Number
				If Previousjob <> "" Then
					Previousjob = Previousjob +"','"+ l_numblot
				Else
					Previousjob = Previousjob +""+ l_numblot
				End If	
				l_numblot = trim(Request.Form("txtLotId"&numblot&""))
				'Get the Job number
				'l_cSQLjob = "select * from fcrefs where rf_ref='" & l_numblot & "' and rf_fh_id ='" & l_jobnumbersscanned & "'"
				l_cSQLjob = "select * from fcrefs where rf_ref='" & l_numblot & "' and rf_fh_id ='" & l_jobnumbersscanned & "' and rf_ref not in ('" & Previousjob & "')"
				Set oRs = oConn.Execute(l_cSQLjob)
				IF oRs.EOF then
					If errorlot <> "" Then
						errorlot = errorlot +","+ l_numblot
					Else
						errorlot = errorlot +""+ l_numblot
					End If	
				END IF	
			End If
		Next
		'Response.End
		If errorlot <> "" Then
			errormessage ="Could not update the status, as the following Lots are not correct (Or) Duplicate Lots: "+" "+errorlot
			Response.Write("<BR>")
			Response.Write("<INPUT name=lblUserID readonly " & _
									"value='" & errormessage & "' " & _
									"style='BACKGROUND-COLOR: red; " & _
									"COLOR: " & l_cColor & "; " & _
									"BORDER-BOTTOM-STYLE: none; " & _
									"BORDER-LEFT-STYLE: none; BORDER-RIGHT-STYLE: none; " & _
									"BORDER-TOP-STYLE: none; HEIGHT: 18px; width:800  '>")
			Response.Write("<BR>")
			Response.Write("<B>Please Re-scan all the Lots</B>")
		ELSE			
			IF l_pickup THEN
			'Response.write "PICK UP!!!!!<BR>"
			'10/19/2005 Span-Renuka modified the code for driver Id.
			l_cSql = "exec pr_dispatch_aironb @p_nfh_pkey = 0, @p_cfhid = '" & l_jobnumbersscanned & "', " & _
					 "@p_npcs = '" & l_cPcs & "', " & _
					 "@p_cDrid = '" & l_cDrid & "', " & _
					 "@p_dTime = '" & NOW & "' "
			Set oRs = oConn.Execute(l_cSql)	
			END IF
			IF l_dropoff THEN
			'Response.write "DROP OFF!!!!<BR>"
			'10/19/2005 Span-Renuka modified the code for driver Id.
			l_cSQL = "EXEC pr_dispatch_airclose " & _
								" @p_cjob = '" & l_jobnumbersscanned & "', " & _
								" @p_dTime = '" & NOW & "', " & _
								" @p_cdr_id = '" & l_cDrid & "', " & _
								" @p_nNumBoxes = '" & l_cPcs & "' "
			Set oRs = oConn.Execute(l_cSql)
			END IF
			'Response.write "<br>!!!!!!!!!!!LocationCode="&LocationCode&"<BR>"
			'Response.write "<br>!!!!!!!!!!!DriverID="&DriverID&"<BR>"			
			%>
			<form method="post" action="DriverifabPhoneEmulator.asp" name="form666">
				<input type="hidden" name="LocationCode" value="<%=LocationCode%>">
				<input type="hidden" name="DriverID" value="<%=DriverID%>">
				<input type="hidden" name="PageStatus" value="loggedin">
			</form>
			<SCRIPT LANGUAGE='JavaScript'> 
			document.forms.form666.submit(); 
			</SCRIPT> 
			
			<%
			''''''''Response.Redirect("ifabackdisplaymessage.asp")	
		END IF																
	END IF
	'Response.write "<br>l_pickup="&l_pickup&"<BR>"
	'Response.write "<br>l_dropoff="&l_dropoff&"<BR>"
	'Response.Write(l_cSQLjob)
	IF l_continue THEN
	%>
	
	<BODY leftMargin=0 TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" OnLoad=document.thisForm.txtLotId1.focus()>
	<FORM ACTION="getjobdetails.asp" method=post name=thisForm>
			<TABLE WIDTH="330" cellpadding="0" cellspacing="5">
			<tr><td align="center" colspan="9"><a href="default.asp" class="mainpagelink">Return Home</a></td></tr>
					<input maxlength='25' size='25' name='txtcaller' id='txtcaller' class='inputgeneral' value='<%=Request.Form("txtcaller")%>' type="hidden">
					<input maxlength='20' name='txtstation' id='txtstation' class='input' size='25' value='<%=Request.Form("txtstation")%>' type="hidden">
					<input maxlength='20' name='txtjobnumber' id='txtjobnumber' class='input' size='25' value='<%=Request.Form("txtjobnumber")%>' type="hidden">
					<input maxlength='40' size='40' name='txtorigin' id='txtorigin' class='inputgeneral' value='<%=l_fl_sf_name%>' Type="hidden">
					<input maxlength='40' size='40' name='txtdestination' id='txtdestination' class='input' size='25' value='<%=l_fl_st_name%>' Type="hidden">
					<input maxlength='40' size='40' name='txtstatus' id='txtstatus' class='input' size='25' value='<%=l_fh_status%>' Type="hidden">
				<TR>
					<TD colspan='2'>
						<div class='purpleseparator'>
							<table width='330' BORDER = '0' id=lottable>
								<tr> 
									<td colspan='6' class='subheader'><img src='../images/transpixel.gif' height='2'></td>
								</tr>
								<tr> 
									<td colspan='4' class='subheader'>Scan Lot Numbers</td>
									<!--<td colspan='2' class='subheader' align='left'><%Response.Write("<input id='addMoreLots' type='button' value='Add More Lots' name='cmdQuote' class='btn' onMouseOver=className='btnh' onMouseOut=className='btn' onclick='generateRow()'>")%></td>-->
									<td colspan='2' class='subheader' align='left'></td>
								</tr>
							<%	
								int numblot
								int in_l_fl_numboxes
								in_l_fl_numboxes = int(l_fl_numboxes)
								For numblot = 1 to in_l_fl_numboxes
							%>
								<!--12/12/2005 Span-Renuka modified the code(Changed maxlength from 8 to 14) for allowing up to 13 characters-->
								<% IF in_l_fl_numboxes >= numblot THEN %>
									<tr><td width='17%'><input id=txtLotId<%=numblot%> maxlength=14 size=20 name=txtLotId<%=numblot%> onchange="validatelot('txtLotId<%=numblot%>')" ></td></tr>
								<% ELSE %>
									<tr><td width='17%'><input id=txtLotId<%=numblot%> maxlength=14 size=20 name=txtLotId<%=numblot%> onchange="validatelot('txtLotId<%=numblot%>')" readonly></td></tr>
								<% 
								END IF
								numblot=numblot+1
								IF in_l_fl_numboxes >= numblot THEN 
								%>
									<tr><td width='17%'><input id=txtLotId<%=numblot%> maxlength=14 size=20 name=txtLotId<%=numblot%> onchange="validatelot('txtLotId<%=numblot%>')"></td></tr>
								<% ELSE %>
									<tr><td width='17%'><input id=txtLotId<%=numblot%> maxlength=14 size=20 name=txtLotId<%=numblot%> onchange="validatelot('txtLotId<%=numblot%>')" readonly></td></tr>
								<% 
								END IF 
								numblot=numblot+1
								IF in_l_fl_numboxes >= numblot THEN 
								%>
									<tr><td width='17%'><input id=txtLotId<%=numblot%> maxlength=14 size=20 name=txtLotId<%=numblot%> onchange="validatelot('txtLotId<%=numblot%>')"></td></tr>
								<% ELSE %>
									<tr><td width='17%'><input id=txtLotId<%=numblot%> maxlength=14 size=20 name=txtLotId<%=numblot%> onchange="validatelot('txtLotId<%=numblot%>')" readonly></td></tr>
								<% 
								END IF 
								numblot=numblot+1
								IF in_l_fl_numboxes >= numblot THEN 
								%>
								<tr><td width='17%'><input id=txtLotId<%=numblot%> maxlength=14 size=20 name=txtLotId<%=numblot%> onchange="validatelot('txtLotId<%=numblot%>')"></td></tr>
								<% ELSE %>
									<tr><td width='17%'><input id=txtLotId<%=numblot%> maxlength=14 size=20 name=txtLotId<%=numblot%> onchange="validatelot('txtLotId<%=numblot%>')" readonly></td></tr>
								<% 
								END IF
								numblot=numblot+1
								IF in_l_fl_numboxes >= numblot THEN
								%>
									<tr><td width='17%'><input id=txtLotId<%=numblot%> maxlength=14 size=20 name=txtLotId<%=numblot%> onchange="validatelot('txtLotId<%=numblot%>')"></td></tr>
								<% ELSE %>
									<tr><td width='17%'><input id=txtLotId<%=numblot%> maxlength=14 size=20 name=txtLotId<%=numblot%> onchange="validatelot('txtLotId<%=numblot%>')" readonly></td></tr>
								<% END IF %>
								<td width='15%'>&nbsp;</td></tr>
							<% 
								Next
							%>
							<input id=txtPieces maxlength=10 size=10 name=txtPieces value=0 type="hidden">&nbsp;</td>
							</table>
						</div>
					</td>
				</TR>
			</TABLE>
			<input id='cmdDriverUpd' type='button' value='Update Job' name='cmdDriverUpd' onclick='updjob()' class='btn' onMouseOver=className='btnh' onMouseOut=className='btn'>
			<input type="hidden" name="LocationCode" value="<%=LocationCode%>" ID="Hidden1">
			<input type="hidden" name="DriverID" value="<%=DriverID%>" ID="Hidden2">
			<INPUT name=lotnumbertext type=hidden value=25>
			<INPUT name=lotnumbersscanned type=hidden value=0>
			<INPUT name=lotnumberstobescanned type=hidden value=<%=l_fl_numboxes%>>
			<INPUT name=savemode type=hidden>
		</FORM>
	</BODY>
	<%
	ELSE
	'Response.Write("Notvalid")
	Response.Redirect("ifabdriverdetails.asp?job=notvalid")
	END IF
	%>
</HTML>
