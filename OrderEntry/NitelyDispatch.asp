<html>
<head>
<!-- #include file="../fleetexpress.inc" -->
<head>
</body>
<% 
sendmessage = "NitelyDispatch.asp started " & Now() & "<br><br>" 
response.write "NitelyDispatch.asp started " & Now() & "<br><br>" 
Set RSEVENTS2 = Server.CreateObject("ADODB.Connection")
RSEVENTS2.ConnectionTimeout = 100
RSEVENTS2.Provider = "MSDASQL"
RSEVENTS2.Open DATABASE

l_cSQL = "SELECT     fcfgthd.fh_id, fcfgthd.fh_user4, fcfgthd.fh_bt_id, fcfgthd.fh_ready, fclegs.fl_sf_name, fclegs.fl_st_name, fclegs.fl_st_rta, fcfgthd.fh_ship_dt,  " &_
     " fcfgthd.fh_carr_id, fcfgthd.fh_status, fcrefs.rf_box  " &_
     "FROM         fcfgthd INNER JOIN " &_
     " fclegs ON fcfgthd.fh_id = fclegs.fl_fh_id INNER JOIN  " &_
     " fcrefs ON fcfgthd.fh_id = fcrefs.rf_fh_id " &_
     " WHERE     (fcfgthd.fh_status = 'RAP' OR fcfgthd.fh_status = 'SCD')"  
                                                  
SET oRs = RSEVENTS2.Execute(l_cSql)
If NOT oRs.eof then
  
             'get new batch nbr
            Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
            
              lSQL = "SELECT MAX(jb_batch_id) as BID FROM JobBatches"                            
							SET oRsB = oConn.Execute(lSQL)
              BID = oRsB("BID")
              if NOT isNumeric(BID) then
                BID = 0
              end if
              BID = cint(BID)
              BID = BID + 1
              
            Set oRsB = Nothing  
						Set oConn=Nothing

            sendmessage = sendmessage & "Batch " & BID & " assigned " & "<br><br>"
            
     Do While not oRs.EOF	
        fh_id=Trim(oRs("fh_id"))
        fh_user4=Trim(oRs("fh_user4"))
        fh_bt_id = Trim(oRs("fh_bt_id"))
                                                                                                           
        response.write "24 order/job=" & fh_id & "<br>"
        JobNum = fh_id
        courier = 0
        if fh_user4 = "Bobtail" then
          courier = "634449"
        elseif fh_user4 = "Van" then
          courier = "307344"
        else
          courier = 0
        end if
        response.write "fh_user4=" & fh_user4 & ", courier=" & courier & "<br>"
        sendmessage = sendmessage & "JobNum " & JobNum & ", Vehicle Type=" & fh_user4 & ", Courier=" & courier & "<br>"
        if courier <> 0 then
                   	 Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
                    		RSEVENTS.CursorLocation = 3
                    		RSEVENTS.CursorType = 3
                   		  RSEVENTS.ActiveConnection = DATABASE
                        SQL = "SELECT fh_status, fh_ship_dt, fh_ready, fh_co_id, fh_co_phone, fh_co_email, fh_co_costcenter, fh_carr_id, fh_custpo, fh_priority, fh_ref FROM fcfgthd where (fh_id = '"& jobnum &"')"
                    		Response.Write "32 SQL="&SQL&"<BR>"
                    		RSEVENTS.Open SQL, DATABASE, 1, 3
                        fh_status=RSEVENTS("fh_status")
                        fh_ship_dt=RSEVENTS("fh_ship_dt")
                        PickUpDateTime=RSEVENTS("fh_ready")
                        RequestorName=RSEVENTS("fh_co_id")
                        Requestorphonenumber=RSEVENTS("fh_co_phone")
                        Requestoremailaddress=RSEVENTS("fh_co_email")
                        if trim(len(requestoremailaddress)) < 1 or requestoremailaddres ="" then
                              requestoremailaddress = "mark.maggiore@logisticorp.us"
                        end if
                        'response.write "614 requestor=" & requestoremailaddress & "<br>"
                        costcenternumber=RSEVENTS("fh_co_costcenter")
                        'IsPalletized=RSEVENTS("IsPalletized")
                        fh_carr_id=RSEVENTS("fh_carr_id")
                        ponumber=RSEVENTS("fh_custpo")
                        priority=RSEVENTS("fh_priority")
                        fh_ref=RSEVENTS("fh_ref")
                        'response.write "fh_ref="&fh_ref&"<BR>"
                    		RSEVENTS.close
                    	  Set RSEVENTS = Nothing
                    	  Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
                    		RSEVENTS.CursorLocation = 3
                    		RSEVENTS.CursorType = 3
                    		RSEVENTS.ActiveConnection = DATABASE
                    		SQL = "SELECT fl_sf_name, fl_sf_clname, fl_sf_phone, fl_sf_email, fl_sf_addr1, fl_sf_city, fl_sf_state, fl_sf_country, fl_sf_zip, fl_sf_comment, fl_st_rta, fl_st_name, fl_st_clname, fl_st_phone, fl_st_email, fl_st_addr1, fl_st_city, fl_st_state, fl_st_country, fl_st_zip, fl_t_release, fl_t_atd   FROM fclegs where (fl_fh_id = '"& jobnum &"')"
                    		'Response.Write "SQL="&SQL&"<BR>"
                    		RSEVENTS.Open SQL, DATABASE, 1, 3
                            OriginationCompany=RSEVENTS("fl_sf_name")
                            OriginationContactname=RSEVENTS("fl_sf_clname")
                            OriginationphoneNumber=RSEVENTS("fl_sf_phone")
                            Originationemail=RSEVENTS("fl_sf_email")
                            Originationaddress=RSEVENTS("fl_sf_addr1")
                            Originationcity=RSEVENTS("fl_sf_city")
                            Originationstate=RSEVENTS("fl_sf_state")
                            Originationcountry=RSEVENTS("fl_sf_country")
                            Originationzip=RSEVENTS("fl_sf_zip")
                            'Pieces=RSEVENTS("NumberOfPieces")
                            comments=RSEVENTS("fl_sf_comment")
                            DeliveryDateTime=RSEVENTS("fl_st_rta")
                            DestinationCompany=RSEVENTS("fl_st_name")
                            DestinationContactname=RSEVENTS("fl_st_clname")
                            DestinationphoneNumber=RSEVENTS("fl_st_phone")
                            Destinationemail=RSEVENTS("fl_st_email")
                            Destinationaddress=RSEVENTS("fl_st_addr1")
                            Destinationcity=RSEVENTS("fl_st_city")
                            Destinationstate=RSEVENTS("fl_st_state")
                            Destinationcountry=RSEVENTS("fl_st_country")
                            Destinationzipcode=RSEVENTS("fl_st_zip")
                            fl_t_release=RSEVENTS("fl_t_release")
                            fl_t_atd=RSEVENTS("fl_t_atd")
                    		RSEVENTS.close
                    	Set RSEVENTS = Nothing
                    	Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
                    		RSEVENTS.CursorLocation = 3
                    		RSEVENTS.CursorType = 3
                    		RSEVENTS.ActiveConnection = DATABASE
                    		SQL = "SELECT rf_box, POD, NumberOfPieces, IsPalletized, Weight, DimLength, DimWidth, DimHeight, Hazmat, Refrigerate FROM fcrefs where (rf_fh_id = '"& jobnum &"')"
                    		'Response.Write "SQL="&SQL&"<BR>"
                    		RSEVENTS.Open SQL, DATABASE, 1, 3
                            rf_box=RSEVENTS("rf_box")
                            POD=RSEVENTS("POD")
                            Pieces=RSEVENTS("NumberOfPieces")
                            IsPalletized=RSEVENTS("IsPalletized")
                            DimWeight=RSEVENTS("Weight")
                            DimLength=RSEVENTS("DimLength")
                            DimWidth=RSEVENTS("DimWidth")
                            DimHeight=RSEVENTS("DimHeight")
                            Hazmat=RSEVENTS("Hazmat")
                            Refrigerate=RSEVENTS("Refrigerate")
                    		RSEVENTS.close
                    	Set RSEVENTS = Nothing                       
	                    Set RSEVENTS = Server.CreateObject("ADODB.Recordset")
		                    RSEVENTS.CursorLocation = 3
		                    RSEVENTS.CursorType = 3
                            'Response.write "Database="&Database&"<br>"
		                    RSEVENTS.ActiveConnection = DATABASE
                            SQL = "SELECT un_id, un_dr_id FROM fcunits where (un_desc = '"& trim(courier) &"')"
		                    'Response.Write "line 778 SQL="&SQL&"<BR>"
                            'Response.Write "DATABASE="&DATABASE&"<BR>"
		                    RSEVENTS.Open SQL, DATABASE, 1, 3
                            un_id=RSEVENTS("un_id")
                            un_dr_id=RSEVENTS("un_dr_id")
                            'response.write "fh_ref="&fh_ref&"<BR>"
		                    RSEVENTS.close
	                    Set RSEVENTS = Nothing


                    DeliveryPeriod=DateDiff("h", PickUpDateTime, DeliveryDateTime)
                    'Response.write "DeliveryPeriod="&DeliveryPeriod&"<br>"
                         Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open INTRANET
							l_cSQL = "UPDATE FleetRouted SET status = 'x' WHERE fh_id = '" & JobNum & "'"
							Response.write "127 l_cSQL="&l_cSQL&"<BR>"
							'oConn.Execute(l_cSQL)
						Set oConn=Nothing
  			            Set RSEVENTS3 = Server.CreateObject("ADODB.Recordset")
				            RSEVENTS3.Open "FleetRouted", Intranet, 2, 2
				            RSEVENTS3.addnew
                            RSEVENTS3("fh_id")=JobNum
                            RSEVENTS3("PickDrop")="p"
                            RSEVENTS3("Courier")=Courier
                            RSEVENTS3("Location")=OriginationCompany
                            RSEVENTS3("ArrivalTime")=PickUpDateTime
                            RSEVENTS3("Pieces")=Pieces
                            RSEVENTS3("BTID")=Session("sBT_ID")
                            RSEVENTS3("DeliveryPeriod")=DeliveryPeriod
                            RSEVENTS3("Status")="c"
				            RSEVENTS3.update
				            'RSEVENTS3.close			
			            'set RSEVENTS3 = nothing                                       
                        
  			            Set RSEVENTS3 = Server.CreateObject("ADODB.Recordset")
				            RSEVENTS3.Open "FleetRouted", Intranet, 2, 2
				            RSEVENTS3.addnew
                            RSEVENTS3("fh_id")=JobNum
                            RSEVENTS3("PickDrop")="d"
                            RSEVENTS3("Courier")=Courier
                            RSEVENTS3("Location")=DestinationCompany
                            RSEVENTS3("ArrivalTime")=DeliveryDateTime
                            'RSEVENTS3("Pieces")=Pieces
                            RSEVENTS3("BTID")=Session("sBT_ID")
                            RSEVENTS3("DeliveryPeriod")=DeliveryPeriod
                            RSEVENTS3("Status")="c"
				            RSEVENTS3.update
				            RSEVENTS3.close			
			            set RSEVENTS3 = nothing  



						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
							l_cSQL = "UPDATE fcfgthd SET fh_status = 'OPN', fh_statcode = 3, fh_dispatcher = '"&UserID&"', fh_ref='"& ReferenceNumber &"', fh_carr_ID='"& Courier &"' WHERE fh_id = '" & JobNum & "'"

                            Response.write "170 l_cSQL="&l_cSQL&"<BR>"

                            oConn.Execute(l_cSQL)
						Set oConn=Nothing
						Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
                       
                        
                        
							l_cSQL = "UPDATE fclegs SET fl_un_id= '"& un_id &"', fl_dr_id='"& un_dr_id &"', fl_t_disp = '"& CurrentDate &"', fl_sf_comment = '"& NewComments &"' WHERE fl_fh_id = '" & JobNum & "'"
                  
                        
                        
                        		
							Response.write "l_cSQL="&l_cSQL&"<BR>"
							oConn.Execute(l_cSQL)
						Set oConn=Nothing

						'update job batches
            Set oConn = Server.CreateObject("ADODB.Connection")
						oConn.ConnectionTimeout = 100
						oConn.Provider = "MSDASQL"
						oConn.Open DATABASE
            
 							l_cSQL = "INSERT INTO JobBatches (jb_batch_id,jb_fh_id) VALUES(" & BID & ",'" & JobNum & "')"
							Response.write "197 l_cSQL="&l_cSQL&"<BR>"
							oConn.Execute(l_cSQL)
						Set oConn=Nothing

      end if
      oRs.movenext
    	LOOP
      
  oRs.close
  Set oRs=Nothing
  RSEVENTS2.Close
  Set RSEVENTS2=Nothing
  response.write "<br>DONE!<br>"
  sendmessage = sendmessage & "<br><br>DONE! " & Now() & "<br><br>"
else
  response.write "No pending Jobs to be dispatched"
  sendmessage = sendmessage & "No pending Jobs to be dispatched " & Now()  & "<br><br>"
end if


'email to Mark and me -

				    Body = sendmessage
				        Set objMail = CreateObject("CDONTS.Newmail")
				        objMail.From = "FleetX@LogisticorpGroup.com"
				        objMail.To = "mark.maggiore@logisticorp.us"
				        objMail.Subject = "Nitely Dispatch Run"
				        objMail.MailFormat = cdoMailFormatMIME
				        objMail.BodyFormat = cdoBodyFormatHTML
				        objMail.Body = Body
				        objMail.Send
                   'End if
				    Set objMail = Nothing
            response.write "<br><br>email sent: " & Body & "<br>"

%>
</body>
</html>