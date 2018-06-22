<table width="100%" cellpadding="0" cellspacing="0" border="0">
    <tr>
        <td><a href="<%=WhichSite%>/Home.asp" class="GreyNavLinks">HOME</a></td>
        <td><a href="<%=WhichSite%>/Tracking/Tracking.asp" class="GreyNavLinks">TRACKING</a></td>
        <td><a href="<%=WhichSite%>/Reporting/default.asp" class="GreyNavLinks">REPORTS</a></td>
        <td>
        <%
        'Response.write "userid="&Userid&"<BR>"
        If userid=1 or userid=221 or trim(lcase(RequestorCompany))="on target" then %>
            <a href="<%=WhichSite%>/OrderEntry/StockRoomOrder.asp" class="GreyNavLinks">
            <%else %>
            <a href="<%=WhichSite%>/OrderEntry/FreightOrder.asp" class="GreyNavLinks">
        <%end if %>
        ORDER ENTRY</a></td>
        <!--
        <td><a href="" class="GreyNavLinks">PHONE EMULATOR</a></td>
        -->
        <td><a href="<%=WhichSite%>/Login.asp?logout=y" class="GreyNavLinks">LOGOUT</a></td>
    </tr> 
</table>
