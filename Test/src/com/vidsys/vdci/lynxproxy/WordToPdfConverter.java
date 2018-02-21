package com.vidsys.vdci.lynxproxy;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import com.itextpdf.text.Chunk;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.Image;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.PageSize;

public class WordToPdfConverter{
 public static void main(String[] args){
   selectFiles();
 }
 public static void selectFiles(){
  JFileChooser chooser = new JFileChooser();
  FileNameExtensionFilter filter = new FileNameExtensionFilter("Microsoft Word 2007+", "docx");
  chooser.setFileFilter(filter);
  chooser.setMultiSelectionEnabled(true);
  int returnVal = chooser.showOpenDialog(null);
  if(returnVal == JFileChooser.APPROVE_OPTION) {
   File[] Files=chooser.getSelectedFiles();
   System.out.println("Please wait...");
           for( int i=0;i<=Files.length;i++){     
            String wordfile=Files[i].toString();            
            convertWordToPdf(wordfile,wordfile.substring(0,wordfile.indexOf('.'))+".pdf");
            }
   System.out.println("Conversion complete");
          }
 }
 public static void convertWordToPdf(String src, String desc){
  try{
  //create file inputstream object to read data from file 
  FileInputStream fs=new FileInputStream(src);
  //create document object to wrap the file inputstream object
  XWPFDocument doc=new XWPFDocument(fs); 
  //72 units=1 inch
  Document pdfdoc=new Document(PageSize.A4,72,72,72,72);
  //create a pdf writer object to write text to mypdf.pdf file
  PdfWriter pwriter=PdfWriter.getInstance(pdfdoc, new FileOutputStream(desc));
  //specify the vertical space between the lines of text
  pwriter.setInitialLeading(20);
  //get all paragraphs from word docx
  List plist=doc.getParagraphs();
  
  //open pdf document for writing
  pdfdoc.open();
  for (int i = 0; i<= plist.size(); i++) {
   //read through the list of paragraphs
   XWPFParagraph pa = plist.get(i);
   //get all run objects from each paragraph
      List runs = pa.getRuns();
      //read through the run objects
      for (int j = 0; j &lt; runs.size(); j++) {       
       XWPFRun run=runs.get(j);
       //get pictures from the run and add them to the pdf document
       List piclist=run.getEmbeddedPictures();
       //traverse through the list and write each image to a file
       Iterator iterator=piclist.iterator();
       while(iterator.hasNext()){
          XWPFPicture pic=iterator.next();
          XWPFPictureData picdata=pic.getPictureData();
          byte[] bytepic=picdata.getData(); 
          Image imag=Image.getInstance(bytepic);
         pdfdoc.add(imag);
             
       }
       //get color code
       int color=getCode(run.getColor());
       //construct font object
       Font f=null;
       if(run.isBold() &amp;&amp; run.isItalic())
        f=FontFactory.getFont(FontFactory.TIMES_ROMAN,run.getFontSize(),Font.BOLDITALIC, new BaseColor(color));
          else if(run.isBold())
        f=FontFactory.getFont(FontFactory.TIMES_ROMAN,run.getFontSize(),Font.BOLD, new BaseColor(color));
       else if(run.isItalic())
        f=FontFactory.getFont(FontFactory.TIMES_ROMAN,run.getFontSize(),Font.ITALIC, new BaseColor(color));
       else if(run.isStrike())
        f=FontFactory.getFont(FontFactory.TIMES_ROMAN,run.getFontSize(),Font.STRIKETHRU, new BaseColor(color));
       else
        f=FontFactory.getFont(FontFactory.TIMES_ROMAN,run.getFontSize(),Font.NORMAL, new BaseColor(color));
       //construct unicode string
       String text=run.getText(-1);
       byte[] bs;
       if (text!=null){
        bs=text.getBytes();
        String str=new String(bs,"UTF-8");
        //add string to the pdf document
        Chunk chObj1=new Chunk(str,f);
        pdfdoc.add(chObj1);
       }      
       
      }
      //output new line
      pdfdoc.add(new Chunk(Chunk.NEWLINE));
      }
     //close pdf document  
  pdfdoc.close();
  }catch(Exception e){e.printStackTrace();}
 } 
 public static int getCode(String code){
  int colorCode;
  if(code!=null)
   colorCode=Long.decode("0x"+code).intValue();
  else
   colorCode=Long.decode("0x000000").intValue();
  return colorCode;
 }

}