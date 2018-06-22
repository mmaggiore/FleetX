// Signature Capture Class 

import java.awt.event.*;
import java.awt.*;
import java.applet.*;
import java.util.*;

public class SigBlock extends Applet{
    SigArea panel;   
    
    public void init() {         
    setLayout(new BorderLayout());    
	panel = new SigArea();
	add("Center", panel);	
}

public void destroy() {
    remove(panel);  
}
    public static void main(String args[]) {
	Frame f = new Frame("SigBlock");
	SigBlock SigBlock = new SigBlock();
	SigBlock.init();	
	f.add("Center", SigBlock);
	f.setSize(300,75);
	f.show();			
	}

public String getAppletInfo() {	    	
    return "Signature capture applet.";
    }

public String sign(){    	
        return panel.sign();
    	}
    	
public void clear(){    	
        panel.clear();
    	}    	

public void lock(){    	
        panel.lock();
    	}    	
    	
public void unlock(){    	
        panel.unlock();
    	}    	

public void decode(String dc){
	    if (dc==null){return;}
	    if (dc!="") {panel.decode(dc);}
    	}
    	
}
class SigArea extends Panel implements MouseListener, MouseMotionListener {    
    Vector lines = new Vector();    
    int x1,y1;
    int x2,y2;    
    int clearflag=0;
    int appwidth=300;
    int appheight=75;
    int delta=15;
    int lock=0;

    public String sign(){    	
    	int np=lines.size();
    	if (np==0){return "";}
    	String temp="";    	
    	for (int i=0; i < np; i++) {
	         Rectangle p = (Rectangle)lines.elementAt(i);
	         temp=temp + p.x + ",";
	         temp=temp + p.y + ",";	         
	         temp=temp + p.width + ",";
	         temp=temp + p.height + ",";
	    }   	    
	    temp = temp.substring(0,temp.length()-1);
	    return temp;
	}		


public void lock(){
	lock=1;
}

public void unlock(){
	lock=0;
}
public void clear(){
	lines.clear();
	repaint();
}

public void decode(String dc){  
        try{                
        String [] coords = dc.split(",");
        int w,x,y,z=0;
        lines.clear();
        repaint();
        for (int i=0; i < coords.length;i=i+4) {              
            w = Integer.parseInt(coords[i]);
            x = Integer.parseInt(coords[i+1]);
            y = Integer.parseInt(coords[i+2]);
            z = Integer.parseInt(coords[i+3]);
            lines.addElement(new Rectangle(w,x,y,z));
        }        
        coords=null;
	    repaint();
	    }	    
	    catch (Exception e)	{}
	    }	    	    

    public SigArea() {
	//setBackground(new Color(255,255,191));			
	setBackground(Color.white);			
	addMouseMotionListener(this);
	addMouseListener(this);	
    }
    public void mouseDragged(MouseEvent e) {
        e.consume();        
        if (lock==1) {return;}
        lines.addElement(new Rectangle(x1, y1, e.getX(), e.getY()));        
        x1 = e.getX();
        y1 = e.getY();         
        repaint();
    }
    public void mouseMoved(MouseEvent e) {    	            
	}
    
    public void mousePressed(MouseEvent e) {
        e.consume();        
        if (lock==1) {return;}
        lines.addElement(new Rectangle(e.getX(), e.getY(), -1, -1));
        x1 = e.getX();
        y1 = e.getY();
        repaint();        
    }
    public void mouseReleased(MouseEvent e) {        
    }
    public void mouseEntered(MouseEvent e) {    	        
    }
    public void mouseExited(MouseEvent e) {
    }
    public void mouseClicked(MouseEvent e) {
    	e.consume();
    if (lock==1) {return;}
    if (e.getX()>3 && e.getX()<13 && e.getY()>3 && e.getY()<13) {		
		clearflag=1;
		repaint();        
		}
    }
        
    public void paint(Graphics g) {
	int np;	
	g.setColor(getBackground());
	if (clearflag==1){
		clearflag=0;
		lines.clear();
	}	
	np = lines.size();			        
	g.setColor(getForeground());
	g.drawLine(0,0,0,delta);
	g.drawLine(0,0,delta,0);
	g.drawLine(appwidth,0,appwidth-delta,0);
	g.drawLine(appwidth,0,appwidth,delta);
	g.drawLine(0,appheight,0,appheight-delta);
	g.drawLine(0,appheight,delta,appheight);
	g.drawLine(appwidth,appheight,appwidth,appheight-delta);
	g.drawLine(appwidth,appheight,appwidth-delta,appheight);	
		
	try {
		Font f = new Font("Monotype Corsiva",Font.ITALIC,64);
	    g.setFont(f);	
	}
	catch (Exception e) {}
	
	g.setColor(new Color(230,230,230));
	g.drawString("Signature",20,52);
	g.setColor(Color.red);
	g.drawRect(3,3,10,10);
	g.drawLine(3,3,13,13);
	g.drawLine(3,13,13,3);
	g.setColor(Color.black);
	
	for (int i=0; i < np; i++) {
	    Rectangle p = (Rectangle)lines.elementAt(i);
	    g.setColor(Color.black);
	    if (p.width != -1) {
		g.drawLine(p.x, p.y, p.width, p.height);
	    } else {
		g.drawLine(p.x, p.y, p.x, p.y);
	    }	     
	}		
}   
}   
