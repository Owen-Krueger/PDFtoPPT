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
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;


/**
 *
 * @author Owen-Krueger
 */
public class ConvertFile {
    
    /**
     * Main method. Run from "convert" button on Main GUI
     * @throws IOException throw if PDF/PPT not selected
     */
    public static void main() throws IOException{
        try{
            //Try to select PDF to open
            File pdf = openPDF();
            //Try to select PPT to save to
            File ppt = selectPPT();
            XMLSlideShow ss = new XMLSlideShow();
            ss.setPageSize(new java.awt.Dimension(720, 405));

            //Convert images from PDF
            ArrayList images = convertImages(readPDF(pdf));

            //Add images to slideshow
            ss = addImages(ss, images);
            
            //Finish outputting file to PowerPoint file
            finishFile(ppt, ss);
        }
        catch(IOException e){
            //If IOException is thrown
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
        //Filter only PDF documents
        fileChooser.setFileFilter(new FileNameExtensionFilter("PDF Documents (.pdf)", "pdf"));
        
        //Choose PDF file to open
        if(fileChooser.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION){
            String filename = fileChooser.getSelectedFile().toString();
            //If file doesn't end in a PDF
            if(!filename.endsWith(".pdf")){
                openPDF();
            }
            return fileChooser.getSelectedFile();
        }
        else {
            //If no PDF selected
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
        //Filter only PowerPoint documments
        fileChooser.setFileFilter(new FileNameExtensionFilter("PowerPoint File (.pptx)", "pptx"));
        
        //Chose .pptx file to output to
        if (fileChooser.showSaveDialog(frame) == JFileChooser.APPROVE_OPTION) {
            String filename = fileChooser.getSelectedFile().toString();
            //If input doesn't end in .pptx
            if(!filename.endsWith(".pptx")){
                filename += ".pptx";
            }
            return new File(filename);
        }
        else {
            //If no output selected
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
        //Load PDF from user chosen file
        pdfDoc = pdfDoc.load(pdfFile);
        ArrayList images = new ArrayList<>();
        
        PDFRenderer pdfRenderer = new PDFRenderer(pdfDoc);
        
        //Render images from file (per page)
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
        
        //Create byte array from each image
        for (BufferedImage image : images) {
            //System.out.println(image.getHeight() + " / " + image.getWidth());
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
        
        //Add pictures to each slide
        for(byte[] picture : pictures){
            slide = ppt.createSlide();
            System.out.println(slide.getSlideLayout());
            pd = ppt.addPicture(picture, XSLFPictureData.PictureType.JPEG);
            ps = slide.createPicture(pd);
        }
        
        return ppt;
    }
    
    /**
     * Finish File output
     * @param outFile File to output to
     * @param ppt Slideshow being outputted to file
     * @throws FileNotFoundException if file not found
     * @throws IOException if input/output issues
     */
    public static void finishFile(File outFile, XMLSlideShow ppt) throws FileNotFoundException, IOException{
        //Output to .pptx file
        FileOutputStream out = new FileOutputStream(outFile);
        ppt.write(out);
    }
}
