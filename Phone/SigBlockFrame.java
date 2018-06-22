import java.awt.event.*;
import java.awt.*;
import java.applet.*;
import java.util.*;
//import java.net.*;
//import javax.swing.*;
//import java.awt.image.*;


/**
 * Sample application using Frame.
 *
 * @author 
 * @version 1.00 05/11/22
 */
public class SigBlockFrame extends Frame {
    
    /**
     * The constructor.
     */  
     public SigBlockFrame() {
                
        MenuBar menuBar = new MenuBar();
        Menu menuFile = new Menu();
        MenuItem menuFileExit = new MenuItem();
        
        menuFile.setLabel("File");
        menuFileExit.setLabel("Exit");
        
        // Add action listener.for the menu button
        menuFileExit.addActionListener
        (
            new ActionListener() {
                public void actionPerformed(ActionEvent e) {
                    SigBlockFrame.this.windowClosed();
                }
            }
        ); 
        menuFile.add(menuFileExit);
        menuBar.add(menuFile);
        
        setTitle("Sigblock");
        setMenuBar(menuBar);
        setSize(new Dimension(400, 400));
        
        // Add window listener.
        this.addWindowListener
        (
            new WindowAdapter() {
                public void windowClosing(WindowEvent e) {
                    SigBlockFrame.this.windowClosed();
                }
            }
        );  
    }
    
    
    /**
     * Shutdown procedure when run as an application.
     */
    protected void windowClosed() {
    	
    	// TODO: Check if it is safe to close the application
    	
        // Exit application.
        System.exit(0);
    }
}
