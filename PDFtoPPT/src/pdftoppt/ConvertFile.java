/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pdftoppt;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;

/**
 *
 * @author owk91
 */
public class ConvertFile {
    
    public File openPDF(){
        JFrame frame = new JFrame();
        JFileChooser fileChooser = new JFileChooser();
        
        if(fileChooser.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION){
            String filename = fileChooser.getSelectedFile().toString();
            if(!filename.endsWith(".pdf")){
                openPDF();
            }
            return fileChooser.getSelectedFile();
        }
        else {
            return null;
        }
    }
    
    public File savePPT(){
        JFrame frame = new JFrame();
        JFileChooser fileChooser = new JFileChooser();
        
        if (fileChooser.showSaveDialog(frame) == JFileChooser.APPROVE_OPTION) {
            String filename = fileChooser.getSelectedFile().toString();
            if(!filename.endsWith(".pptx")){
                filename += ".pptx";
            }
            return new File(filename);
        }
        else {
            return null;
        }
    }
    
    public ArrayList<BufferedImage> readPDF(File pdfFile) throws IOException{
        PDDocument pdfDoc = new PDDocument();
        pdfDoc = pdfDoc.load(pdfFile);
        ArrayList images = new ArrayList<BufferedImage>();
        
        PDFRenderer pdfRenderer = new PDFRenderer(pdfDoc);
        
        for(int i = 0; i < pdfDoc.getNumberOfPages(); i++){
            images.add(pdfRenderer.renderImage(i));
        }
        return images;
    }
    
    public void convertImages(ArrayList<BufferedImage> images){
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        
        for(int i = 0; i < images.size(); i++){
            
        }
    }
}
