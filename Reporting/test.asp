
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Pickadate Example</title>


<link rel="stylesheet" href="../themes/classic.css" id="theme_base">
<link rel="stylesheet" href="../themes/classic.date.css" id="theme_date">
<link rel="stylesheet" href="../themes/classic.time.css" id="theme_time">


</head>
<body>
<%
TheHours=Hour(Time())
TheMinutes=Minute(Time())
Response.write "TheMinutes="&TheMinutes&"<BR>"
If Int(TheHours)>12 then
    Response.write "Got here!<BR>"
    TheHours=TheHours-12
End if
AMPM=Right(Time(),2)
TheTime=TheHours&":"&TheMinutes&" "&AMPM
 %>
	<form>
		<h2>Date Example1</h2>
		<label for="date_1">Field: Date_1</label>
		<input type="text" name="date_1" id="date_1" value="<%=date()%>">
		<span class="offscreen">(Format: mm/dd/yyyy)</span>
	</form>

	<form>
		<h2>Time Picker</h2>
		<label for="time_1">Field: Time_1</label>
		<input type="text" name="time_1" id="time_1" value="<%=TheTime %>">
	</form>


<script src="../jquery-2.1.0.min.js"></script> 
<script src="../pickadate.js"></script> 
<script type="text/javascript">
    // PICKADATE FORMATTING
    $('#date_1').pickadate({
        format: 'mm/dd/yyyy', 	// Friendly format displayed to user
        formatSubmit: 'mm/dd/yyyy', // Actual format used by application
        hiddenName: false			// Allows two different formats
    });


    $('#time_1').pickatime({
        format: 'h:i A', 		// Displayed and application format
        interval: 10, 			// Interval between values (in minutes)
        min: '12:00 AM', 			// Starting value
        max: '11:59 PM'				// Ending value
    });

</script>
</body>
</html>
