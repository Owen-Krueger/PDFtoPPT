package pdftoppt;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
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
public class ConvertFile {
    
    /**
     * Prompts user to select PDF file to open
     * @return File pdf file
     */
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
    
    /**
     * Prompts user with save location for .pptx file
     * @return File PowerPoint file
     */
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
    
    /**
     * Reads PDF file from user and creates ArrayList of images
     * @param pdfFile File pdf from user prompt
     * @return ArrayList<BufferedImage> image of each PDF page
     * @throws IOException throws IO exception if fails to load pdfFile
     */
    public ArrayList<BufferedImage> readPDF(File pdfFile) throws IOException{
        PDDocument pdfDoc = new PDDocument();
        pdfDoc = pdfDoc.load(pdfFile);
        ArrayList images = new ArrayList<>();
        
        PDFRenderer pdfRenderer = new PDFRenderer(pdfDoc);
        
        for(int i = 0; i < pdfDoc.getNumberOfPages(); i++){
            images.add(pdfRenderer.renderImage(i));
        }
        return images;
    }
    
    /**
     * Converts ArrayList of images into Bytes to be used by XMLSlideshow
     * @param images ArrayList<BufferedImage> array of images from PDF
     * @return ArrayList<byte[]> array of converted images
     * @throws IOException throws IO exception if fails to write image to ByteArray
     */
    public ArrayList<byte[]> convertImages(ArrayList<BufferedImage> images) throws IOException{
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ArrayList<byte[]> pictures = new ArrayList<>();
        
        for (BufferedImage image : images) {
            ImageIO.write(image, "jpg", baos);
            byte[] imageByte = baos.toByteArray();
            pictures.add(imageByte);
        }
        
        return pictures;
    }
    
    /**
     * 
     * @param ppt
     * @param pictures
     * @return 
     */
    public XMLSlideShow addImages(XMLSlideShow ppt, ArrayList<byte[]> pictures){
        PictureData pd;
        XSLFSlide slide;
        XSLFPictureShape ps;
        
        for(byte[] picture : pictures){
            pd = ppt.addPicture(picture, XSLFPictureData.PictureType.JPEG);
            slide = ppt.createSlide();
            ps = slide.createPicture(pd);
        }
        
        return ppt;
    }
    
    /**
     * 
     * @param outFile
     * @param ppt
     * @throws FileNotFoundException
     * @throws IOException 
     */
    public void finishFile(File outFile, XMLSlideShow ppt) throws FileNotFoundException, IOException{
        FileOutputStream out = new FileOutputStream(outFile);
        ppt.write(out);
    }
}
