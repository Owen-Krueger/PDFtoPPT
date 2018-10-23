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
    
    public static void main() throws IOException{
        try{
            File pdf = openPDF();
            File ppt = selectPPT();
            XMLSlideShow ss = new XMLSlideShow();

            ArrayList images = convertImages(readPDF(pdf));

            ss = addImages(ss, images);
            finishFile(ppt, ss);
        }
        catch(IOException e){
            throw new IOException(e.getMessage());
        }
    }
    
    /**
     * Prompts user to select PDF file to open
     * @return File pdf file
     * @throws java.io.IOException
     */
    public static File openPDF() throws IOException{
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
            throw new IOException("No PDF Selected");
        }
    }
    
    /**
     * Prompts user with save location for .pptx file
     * @return File PowerPoint file
     * @throws java.io.IOException
     */
    public static File selectPPT() throws IOException{
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
            throw new IOException("No Output Selected");
        }
    }
    
    /**
     * Reads PDF file from user and creates ArrayList of images
     * @param pdfFile File pdf from user prompt
     * @return ArrayList<BufferedImage> image of each PDF page
     * @throws IOException throws IO exception if fails to load pdfFile
     */
    public static ArrayList<BufferedImage> readPDF(File pdfFile) throws IOException{
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
    public static ArrayList<byte[]> convertImages(ArrayList<BufferedImage> images) throws IOException{
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ArrayList<byte[]> pictures = new ArrayList<>();
        
        for (BufferedImage image : images) {
            ImageIO.write(image, "jpg", baos);
            byte[] imageByte = baos.toByteArray();
            pictures.add(imageByte);
            baos.flush();
            baos.reset();
        }
        
        return pictures;
    }
    
    /**
     * Added images to PowerPoint slides
     * @param ppt PowerPoint presentation to add pictures to
     * @param pictures Array of pictures to add to presentation
     * @return PowerPoint presentation
     */
    public static XMLSlideShow addImages(XMLSlideShow ppt, ArrayList<byte[]> pictures){
        PictureData pd;
        XSLFSlide slide;
        XSLFPictureShape ps;
        
        for(byte[] picture : pictures){
            slide = ppt.createSlide();
            pd = ppt.addPicture(picture, XSLFPictureData.PictureType.JPEG);
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
    public static void finishFile(File outFile, XMLSlideShow ppt) throws FileNotFoundException, IOException{
        FileOutputStream out = new FileOutputStream(outFile);
        ppt.write(out);
    }
}
