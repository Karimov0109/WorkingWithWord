//import org.apache.poi.xwpf.usermodel.*;
//import org.apache.xmlbeans.XmlCursor;
//import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
//import org.w3c.dom.Node;
//
//import javax.xml.transform.Transformer;
//import javax.xml.transform.TransformerFactory;
//import javax.xml.transform.dom.DOMSource;
//import javax.xml.transform.stream.StreamResult;
//import javax.xml.transform.stream.StreamSource;
//import java.awt.*;
//import java.io.*;
//
//public class Original{
//
//    static File stylesheet = new File("C:\\Program Files\\Microsoft Office\\root\\Office16\\OMML2MML.XSL");
//static TransformerFactory tFactory = TransformerFactory.newInstance();
//static StreamSource stylesource = new StreamSource(stylesheet);
//
////method for getting MathML from oMath
//static String getMathML(CTOMath ctomath) throws Exception {
//        Transformer transformer = tFactory.newTransformer(stylesource);
//
//        Node node = ctomath.getDomNode();
//
//        DOMSource source = new DOMSource(node);
//        StringWriter stringwriter = new StringWriter();
//        StreamResult result = new StreamResult(stringwriter);
//        transformer.setOutputProperty("omit-xml-declaration", "yes");
//        transformer.transform(source, result);
//
//        String mathML = stringwriter.toString();
//        stringwriter.close();
//
//        //The native OMML2MML.XSL transforms OMML into MathML as XML having special name spaces.
//        //We don't need this since we want using the MathML in HTML, not in XML.
//        //So ideally we should changing the OMML2MML.XSL to not do so.
//        //But to take this example as simple as possible, we are using replace to get rid of the XML specialities.
//        mathML = mathML.replaceAll("xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"", "");
//        mathML = mathML.replaceAll("xmlns:mml", "xmlns");
//        mathML = mathML.replaceAll("mml:", "");
//
//        return mathML;
//        }
//
////method for getting HTML including MathML from XWPFParagraph
//static String getTextAndFormulas(XWPFParagraph paragraph) throws Exception {
//
//        StringBuilder textWithFormulas = new StringBuilder();
//
//        //using a cursor to go through the paragraph from top to down
//        XmlCursor xmlcursor = paragraph.getCTP().newCursor();
//
//        while (xmlcursor.hasNextToken()) {
//        XmlCursor.TokenType tokentype = xmlcursor.toNextToken();
//        if (tokentype.isStart()) {
//        if (xmlcursor.getName().getPrefix().equalsIgnoreCase("w") && xmlcursor.getName().getLocalPart().equalsIgnoreCase("r")) {
//        //elements w:r are text runs within the paragraph
//        //simply append the text data
//        textWithFormulas.append(xmlcursor.getTextValue());
//        } else if (xmlcursor.getName().getLocalPart().equalsIgnoreCase("oMath")) {
//        //we have oMath
//        //append the oMath as MathML
//        textWithFormulas.append(getMathML((CTOMath)xmlcursor.getObject()));
//        }
//        } else if (tokentype.isEnd()) {
//        //we have to check whether we are at the end of the paragraph
//        xmlcursor.push();
//        xmlcursor.toParent();
//        if (xmlcursor.getName().getLocalPart().equalsIgnoreCase("p")) {
//        break;
//        }
//        xmlcursor.pop();
//        }
//        }
//
//        return textWithFormulas.toString();
//        }
//
//public static void main(String[] args) throws Exception {
//
//        XWPFDocument document = new XWPFDocument(new FileInputStream("C:\\Users\\User\\Desktop\\Matematika_baza.docx"));
//
//        //using a StringBuffer for appending all the content as HTML
//        StringBuilder allHTML = new StringBuilder();
//
//        //loop over all IBodyElements - should be self explained
//        for (IBodyElement ibodyelement : document.getBodyElements()) {
//        if (ibodyelement.getElementType().equals(BodyElementType.PARAGRAPH)) {
//        XWPFParagraph paragraph = (XWPFParagraph)ibodyelement;
//        allHTML.append("<p>");
//        allHTML.append(getTextAndFormulas(paragraph));
//        allHTML.append("</p>");
//        } else if (ibodyelement.getElementType().equals(BodyElementType.TABLE)) {
//        XWPFTable table = (XWPFTable)ibodyelement;
//        allHTML.append("<table border=1>");
//        for (XWPFTableRow row : table.getRows()) {
//        allHTML.append("<tr>");
//        for (XWPFTableCell cell : row.getTableCells()) {
//        allHTML.append("<td>");
//        for (XWPFParagraph paragraph : cell.getParagraphs()) {
//        allHTML.append("<p>");
//        allHTML.append(getTextAndFormulas(paragraph));
//        allHTML.append("</p>");
//        }
//        allHTML.append("</td>");
//        }
//        allHTML.append("</tr>");
//        }
//        allHTML.append("</table>");
//        }
//        }
//
//        document.close();
//
//        //creating a sample HTML file
//        String encoding = "UTF-8";
//        FileOutputStream fos = new FileOutputStream("C:\\Users\\User\\Desktop\\result.html");
//        OutputStreamWriter writer = new OutputStreamWriter(fos, encoding);
//        writer.write("<!DOCTYPE html>\n");
//        writer.write("<html lang=\"en\">");
//        writer.write("<head>");
//        writer.write("<meta charset=\"utf-8\"/>");
//
//        //using MathJax for helping all browsers to interpret MathML
//        writer.write("<script type=\"text/javascript\"");
//        writer.write(" async src=\"https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.1/MathJax.js?config=MML_CHTML\"");
//        writer.write(">");
//        writer.write("</script>");
//
//        writer.write("</head>");
//        writer.write("<body>");
//
//        writer.write(allHTML.toString());
//
//        writer.write("</body>");
//        writer.write("</html>");
//        writer.close();
//
//        Desktop.getDesktop().browse(new File("C:\\Users\\User\\Desktop\\result.html").toURI());
//
//        }}