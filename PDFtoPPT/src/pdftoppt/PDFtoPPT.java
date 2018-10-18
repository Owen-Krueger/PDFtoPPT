/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pdftoppt;

import java.awt.image.BufferedImage;
import java.io.*;
import javax.imageio.ImageIO;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;


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
        BufferedImage image = null;
        
        if(fileChooser.showOpenDialog(parent) == JFileChooser.APPROVE_OPTION){
            pdfFile = fileChooser.getSelectedFile();
            System.out.println(pdfFile);
            pdf = pdf.load(pdfFile);
            System.out.println(pdf.getPage(0).toString());
            pdfRender = new PDFRenderer(pdf);
            System.out.println(pdfRender.toString());
            //for(int i = 0; i < pdf.getNumberOfPages(); i++)
            image = pdfRender.renderImage(0);
            
            
        }
        
        if (fileChooser.showSaveDialog(parent) == JFileChooser.APPROVE_OPTION) {
            String filename = fileChooser.getSelectedFile().toString();
            if(!filename.endsWith(".pptx")){
                filename += ".pptx";
            }
            
            File file = new File(filename);
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ImageIO.write( image, "jpg", baos );
            byte[] imageByte = baos.toByteArray();
            //File file = fileChooser.getSelectedFile();
            
            PictureData idx = ppt.addPicture(imageByte, XSLFPictureData.PictureType.JPEG);
      
            XSLFSlide slide = ppt.createSlide();
            
      //creating a slide with given picture on it
            XSLFPictureShape pic = slide.createPicture(idx);
            
            FileOutputStream out = new FileOutputStream(file);
            
            //ppt.addPicture(image, PictureData.PictureType.EMF);
            ppt.write(out);            
            
        }
        System.exit(0);
        System.out.println("Done");
    }    
}
