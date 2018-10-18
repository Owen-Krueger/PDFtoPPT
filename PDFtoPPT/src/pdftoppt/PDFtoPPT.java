/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pdftoppt;

import java.io.*;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import org.apache.poi.xslf.usermodel.XMLSlideShow;


/**
 *
 * @author owk91
 */
public class PDFtoPPT {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws FileNotFoundException, IOException{
        JFrame parent = new JFrame();
        XMLSlideShow ppt = new XMLSlideShow();
         
        JFileChooser fileChooser = new JFileChooser();
        if (fileChooser.showSaveDialog(parent) == JFileChooser.APPROVE_OPTION) {
            String filename = fileChooser.getSelectedFile().toString();
            if(!filename.endsWith(".pptx")){
                filename += ".pptx";
            }
            
            File file = new File(filename);
            //File file = fileChooser.getSelectedFile();
            FileOutputStream out = new FileOutputStream(file);
            ppt.createSlide();
            ppt.write(out);            
            
        }
        System.exit(0);
        System.out.println("Done");
    }    
}
