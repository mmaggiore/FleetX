     FREE (the best price!) SIGNATURE APPLET
     ---------------------------------------



     ONE LINER - UNPACK TO A DIRECTORY, OPEN EXAMPLE1.HTML
     -----------------------------------------------------


FILES
-----

     This archive should contain the following files:-
   
     1). README.txt - That's this file!
     2). SigBlock.java & SigBlockFrame.java for loading into the free and very excellent 
         JCreator editor found at www.jcreator.com.
     3). SigArea.class, SigBlock.class & SigBlockFrame.class - these are the applet classes
         that you will want to embed in your page. Only embed SigBlock.class but it calls 
         the other two so they need to be in the same directory as the applet. Maybe they
	 need to be in the same directory as the page - I've not tried changing that yet.
     4). example1.html - an example file that shows how to embed and control the applet.
	
     Simply unpack all the files to a directory and open the example1.html with IE or Firefox.


BRIEF
-----

     After spending 2 days getting the damn applet to work (I never used java before!), I decided
     that I would stick it on the web for other people to use. The reasons for this were twofold:-

     FIRSTLY -  I was amazed that such a basic thing was not freely available. There were loads
     of companies quite happy to take money off you in exchange for an applet that may or may not
     do what you want them to... but you can't get to the source code and you can't change stuff
     at the grass roots level.
     
     SECONDLY - Having spent 2 days getting it going I hope that others will not have to waste
     time doing the same thing. Instead you can spend the time customising it to your needs!!!
     I'm not a java programmer so it should probably have taken an hour buy hey!
     
     ACKNOWLEDGEMENTS - I must admit that the basic "draw stuff when the mouse is moved" code
     came from the demo DrawTest applet that ships with JRE. I just hacked out the bits I didn't
     want like colours and the tool panel until I had a basic "Sign here" type applet.
     
     WHAT THE APPLET GIVES YOU - Well you don't get a bitmap or gif out but you do get a list of
     vector coordinates that describe the signature. You can stick this string of digits in a 
     database or file and use the sign() and decode(...) functions of the applet to get and put
     information to the "Sign Here" area.
     
     You can clear the applet by using the unremarkably named clear() method! or by doing a decode()
     with an empty string. Other functionality can easily be added as required - I added a lock() and
     unlock() facility so users couldn't change a signature once its been input. 

     FREEWARE - Feel free to use, modify and generally have fun with this applet. Feel free to link back
     to my site to ensure you get any extra functionality I add in the future.
     
     WEB - www.clayzer.com  then check out the Resources/Freeware Section.
     
     BUGS - Please not that although the code works fine in Explorer, there is a slight delay on
     execution of the first function when used in Firefox. This is something to do with the java 
     execution rights. Any ideas on how to prevent this delay would be great!

     ---------------------------------------------------------------------------------

     USAGE - Embed the applet in a page as shown below. If you stick it in a table it makes it 
     much easier to position with align="center" and valign="middle" etc.
     
     The demo below will give you access to most of the features for trying it out.


Remember to visit www.clayzer.com  then Resources/Freeware for any enhancements or updates that come along.

*** ENDS ***