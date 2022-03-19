import org.w3c.dom.Node;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.stream.StreamSource;
import java.io.File;

//import org.apache.poi.openxml4j.opc.OPCPackage;
//import org.apache.poi.xwpf.usermodel.XWPFDocument;
//import org.apache.poi.xwpf.usermodel.XWPFParagraph;
//
//import java.io.FileInputStream;
//
//public class WordGenerator {
//    public static void main(String[] args) {
//        try(FileInputStream fis = new FileInputStream("C:\\Users\\User\\Desktop\\Matematika_baza.docx")) {
//            XWPFDocument doc    = new XWPFDocument(OPCPackage.open(fis));
//            java.util.List<XWPFParagraph> paragraphs =  doc.getParagraphs();
//            for (XWPFParagraph paragraph: paragraphs){
//                System.out.println(paragraph.getParagraphText());
//            }
//        }catch(Exception e) {
//            System.out.println(e);
////        }
////    }
////}
//public  class WordGenerator{
//static File stylesheet =new File("C:\\Users\\User\\Desktop\\WordExcel.xsl");
//static TransformerFactory tFactory = TransformerFactory.newInstance();
//static StreamSource stylesource = new StreamSource(stylesheet);
//static String geyMathML(CTOMath ctoMath)throws Exception{
//    Transformer transformer = tFactory.newTransformer(stylesource);
//    Node node =ctoMath.getDomNode();
//
//}
//}