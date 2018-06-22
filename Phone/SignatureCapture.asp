<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title></title>
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
</head>
<body><!--  OnPageLoad checks for proper function  OnLoad= "OnPageLoad()" 
               of OCX when page is loaded 
          -->
<H2>Bennet-Tec Web Signature Control </H2>
<UL>
  <P>This sample application captures a handwritten <BR>signature within a web 
  page and saves the image <BR>to a local GIF image file <BR>
  <P>Put your signature in the box below <BR>and click the Submit button. 
  <BR></P>
  <P><!-- Embed the Signature Control  -->
  <OBJECT id=WebSignature codeBase=WebSign.i386.CAB#version=1,0,3 height=55 
  onerror=OnOCXError() width=180 
  classid=clsid:4E534257-3000-4E22-83A5-AD614CB96642 name=WebSignature VIEWASTEXT>
	<PARAM NAME="_cx" VALUE="4763">
	<PARAM NAME="_cy" VALUE="1455">
	<PARAM NAME="DrawWidth" VALUE="1">
	<PARAM NAME="BorderWidth" VALUE="-1">
	<PARAM NAME="BaseLine" VALUE="80">
	<PARAM NAME="BaseLineWidth" VALUE="-2">
	<PARAM NAME="ImageWidth" VALUE="0">
	<PARAM NAME="ImageHeight" VALUE="0">
	<PARAM NAME="DrawLineSize" VALUE="2">
	<PARAM NAME="BackColor" VALUE="4294967295">
	<PARAM NAME="BorderColor" VALUE="2147483654">
	<PARAM NAME="BaseLineColor" VALUE="8421504">
	<PARAM NAME="DrawColor" VALUE="0">
	<PARAM NAME="BaseSite" VALUE="">
	<PARAM NAME="LicensedSite" VALUE="DEMO">
	<PARAM NAME="LicenseKey" VALUE="DEMO">
	<PARAM NAME="ReadOnly" VALUE="0">
	<PARAM NAME="TabStop" VALUE="-1">
	</OBJECT>
	
  <P><!-- button to allow user to clear --><INPUT id=Button1 onclick=OnClear() type=button value=Clear name=Button1 Width="400"> 
<!-- button to allow user to Save --><INPUT id=Button2 onclick=OnSubmit() type=button value="Save Signature" name=Button2> 

  <P>Your signature will be saved on your computer <BR>at C:\Signature.Gif 
</P></UL><!-- ===========END OF HTML BODY ============= -->
<SCRIPT language=JScript>


  	/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */
  	/*  - On PageLoad we call this function as a test  - - - - - - */
        /*  - to verify control is properly loaded and working   - - - */
  	/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */
  	function OnPageLoad ()
  	{   /* This routine called by BODY OnLoad event */
	  try
	  {
            /* Call Clear method at start to test 
            /* that the signature control is installed */
	    WebSignature.Clear();
	  }
	  catch(e)
	  {     
            /* If clear does not work then control is not installed */
            /* Redirect user to page with instructions to install */
	    document.location.href = "InstallError.htm";
	  }
 	}


      /* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */
      function OnOCXError ()
      {
	  document.location.href = "OCXerror.htm";
      }


      /* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */
      function OnClear ()
      {
	WebSignature.Clear();
      }

      var WSF_PICTURE = 0x80;		// - Raster Bitmap image
      var WSF_PICTUREGIF = 0xC0;	// - GIF image
      var WSF_ORIGINALSIZE = 1;		// - Use original picture size 
                                	// (do not strip spaces around the signature).

      /* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */
      function OnSubmit ()
      {
	WebSignature.SavePicture("C:\\signature.gif", 
                   WSF_PICTUREGIF + WSF_ORIGINALSIZE ); 

      }
      /* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */

    </SCRIPT>



</body>
</html>
