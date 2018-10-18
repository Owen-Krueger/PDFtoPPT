/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pdftoppt;

import java.awt.image.BufferedImage;
import java.io.*;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.sl.usermodel.PictureData;
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
        PDDocument pdf = new PDDocument();
        PDFRenderer pdfRender;
        JFileChooser fileChooser = new JFileChooser();
        File pdfFile;
        BufferedImage image;
        
        if(fileChooser.showOpenDialog(parent) == JFileChooser.APPROVE_OPTION){
            pdfFile = fileChooser.getSelectedFile();
            pdf.load(pdfFile);
            pdfRender = new PDFRenderer(pdf);
            image = pdfRender.renderImage(0);
            
        }
        
        if (fileChooser.showSaveDialog(parent) == JFileChooser.APPROVE_OPTION) {
            String filename = fileChooser.getSelectedFile().toString();
            if(!filename.endsWith(".pptx")){
                filename += ".pptx";
            }
            
            File file = new File(filename);
            
            //File file = fileChooser.getSelectedFile();
            FileOutputStream out = new FileOutputStream(file);
            ppt.createSlide();
            
            //ppt.addPicture(image, PictureData.PictureType.EMF);
            ppt.write(out);            
            
        }
        System.exit(0);
        System.out.println("Done");
    }    
}
