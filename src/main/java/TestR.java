import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.w3c.dom.Node;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import java.awt.*;
import java.io.*;
import java.util.*;
import java.util.List;
import java.util.regex.Pattern;


public class TestR {

    static File stylesheet = new File("C:\\Program Files\\Microsoft Office\\root\\Office16\\OMML2MML.XSL");
    static TransformerFactory tFactory = TransformerFactory.newInstance();
    static StreamSource stylesource = new StreamSource(stylesheet);

    //method for getting MathML from oMath
    static String getMathML(CTOMath ctomath) throws Exception {
        Transformer transformer = tFactory.newTransformer(stylesource);

        Node node = ctomath.getDomNode();

        DOMSource source = new DOMSource(node);
        StringWriter stringwriter = new StringWriter();
        StreamResult result = new StreamResult(stringwriter);
        transformer.setOutputProperty("omit-xml-declaration", "yes");
        transformer.transform(source, result);

        String mathML = stringwriter.toString();
        stringwriter.close();

        //The native OMML2MML.XSL transforms OMML into MathML as XML having special name spaces.
        //We don't need this since we want using the MathML in HTML, not in XML.
        //So ideally we should changing the OMML2MML.XSL to not do so.
        //But to take this example as simple as possible, we are using replace to get rid of the XML specialities.
        mathML = mathML.replaceAll("xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"", "");
        mathML = mathML.replaceAll("xmlns:mml", "xmlns");
        mathML = mathML.replaceAll("mml:", "");

        return mathML;
    }

    //method for getting HTML including MathML from XWPFParagraph
    static String getTextAndFormulas(XWPFParagraph paragraph) throws Exception {

        StringBuilder textWithFormulas = new StringBuilder();

        //using a cursor to go through the paragraph from top to down
        XmlCursor xmlcursor = paragraph.getCTP().newCursor();

        while (xmlcursor.hasNextToken()) {
            XmlCursor.TokenType tokentype = xmlcursor.toNextToken();
            if (tokentype.isStart()) {
                if (xmlcursor.getName().getPrefix().equalsIgnoreCase("w") && xmlcursor.getName().getLocalPart().equalsIgnoreCase("r")) {
                    //elements w:r are text runs within the paragraph
                    //simply append the text data
                    textWithFormulas.append(xmlcursor.getTextValue());
                } else if (xmlcursor.getName().getLocalPart().equalsIgnoreCase("oMath")) {
                    //we have oMath
                    //append the oMath as MathML
                    textWithFormulas.append(getMathML((CTOMath) xmlcursor.getObject()));
                }
            } else if (tokentype.isEnd()) {
                //we have to check whether we are at the end of the paragraph
                xmlcursor.push();
                xmlcursor.toParent();
                if (xmlcursor.getName().getLocalPart().equalsIgnoreCase("p")) {
                    break;
                }
                xmlcursor.pop();
            }
        }

        return textWithFormulas.toString();
    }


    public static double generateRandom(Integer min) {
        return ((int) Math.floor(Math.random() * (50)) + (min));
    }

//    public static void main(String[] args) throws Exception {
//        readDocxFile("C:\\Users\\User\\Desktop\\Matematika_baza.docx");
//    }

    public static void readDocxFile(String fileName) throws Exception {
        try {
            File file = new File(fileName);
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
            StringBuilder fullText = new StringBuilder();
            StringBuilder allHTML = new StringBuilder();

            XWPFDocument document = new XWPFDocument(fis);
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for (XWPFParagraph paragraph : paragraphs) {
                fullText.append(paragraph.getText()).append("\n");
                allHTML.append("<p>");
                allHTML.append(getTextAndFormulas(paragraph));
                allHTML.append("</p>");
            }


            String[] data = allHTML.toString().split("#");
            List<String> keys = new ArrayList<>();
            List<String> values = new ArrayList<>();
            for (int i = 0; i < data.length - 1; i++) {
                String[] datss = data[i].split("\n");
                String[] answers = datss[0].split("\\.");
                if (answers.length == 2) {
                    keys.add(answers[1]);
                }
                StringBuilder questions = new StringBuilder();

                for (int i1 = 1; i1 < datss.length; i1++) {
                    questions.append(datss[i1]);
                }

                if (questions.length() > 1) {
                    values.add(questions.toString());

                }
            }

            for (String value : values) {
                System.out.println("paragraph -> " + value);
            }

            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static String parseText(String text) {
        CharSequence inputStr = "annb";

        String pattern = "(?<=(rn|r|n))([ \t]*$)+";

        String[] par = Pattern.compile(pattern, Pattern.MULTILINE).split(inputStr);

        for (int i = 0; i < par.length; i++) {
            String paragraph = par[i];
            return paragraph;
        }
        return "";
    }

    public static void WritingMethod() throws IOException {
        String folder = "C:\\Users\\User\\Desktop\\";
        String fileName1 = "PrimaryWord1.docx";


        XWPFDocument document1 = new XWPFDocument();

        //Write the Document in file system
        FileOutputStream out = new FileOutputStream(new File(folder + fileName1));

        while (true) {
            //create Paragraph
            XWPFParagraph paragraph1 = document1.createParagraph();

            XWPFRun run = paragraph1.createRun();
//            run.setText();

            document1.write(out);
            out.close();
            System.out.println("createparagraph.docx written successfully");
        }
    }

    public static void RandomMethod(int yourInput, int counter) {

        List<Integer> random1List = new ArrayList<>();
        List<Integer> random2List = new ArrayList<>();
        List<Integer> random3List = new ArrayList<>();
        List<Integer> random4List = new ArrayList<>();
        List<Integer> random5List = new ArrayList<>();
        List<Integer> random6List = new ArrayList<>();

        while (counter < yourInput) {
            int random1 = (int) generateRandom(1);
            for (Integer random : random1List) {
                if (random == random1) {
                    random1 = (int) generateRandom(1);
                }
            }

            random1List.add(random1);
            counter++;
            if (counter == yourInput) break;


            int random2 = (int) generateRandom(51);
            for (Integer random : random2List) {
                if (random == random2) {
                    random2 = (int) generateRandom(51);
                }
            }
            random2List.add(random2);
            counter++;
            if (counter == yourInput) break;


            int random3 = (int) generateRandom(101);
            for (Integer random : random3List) {
                if (random == random3) {
                    random3 = (int) generateRandom(101);
                }
            }
            random3List.add(random3);
            counter++;
            if (counter == yourInput) break;

            int random4 = (int) generateRandom(151);
            for (Integer random : random4List) {
                if (random == random4) {
                    random4 = (int) generateRandom(151);
                }
            }

            random4List.add(random4);
            counter++;
            if (counter == yourInput) break;


            int random5 = (int) generateRandom(201);
            for (Integer random : random5List) {
                if (random == random2) {
                    random5 = (int) generateRandom(201);
                }
            }
            random5List.add(random5);
            counter++;
            if (counter == yourInput) break;


            int random6 = (int) generateRandom(251);
            for (Integer random : random6List) {
                if (random == random6) {
                    random6 = (int) generateRandom(251);
                }
            }
            random6List.add(random6);
            counter++;
            if (counter == yourInput) break;
        }

        System.out.println("random list1 ->" + random1List);
        System.out.println("random list2 ->" + random2List);
        System.out.println("random list3 ->" + random3List);
        System.out.println("random list4 ->" + random4List);
        System.out.println("random list5 ->" + random5List);
        System.out.println("random list6 ->" + random6List);
    }

    public static void main(String[] args) throws Exception  {

        XWPFDocument document = new XWPFDocument(new FileInputStream("C:\\Users\\User\\Desktop\\Matematika_baza.docx"));
        StringBuilder allHTML = new StringBuilder();
        StringBuilder fullText = new StringBuilder();


        for (IBodyElement ibodyelement : document.getBodyElements()) {
            if (ibodyelement.getElementType().equals(BodyElementType.PARAGRAPH)) {
                XWPFParagraph paragraph = (XWPFParagraph) ibodyelement;

                fullText.append(paragraph.getText()).append("\n");

                allHTML.append("<p>");
                allHTML.append(getTextAndFormulas(paragraph));

                allHTML.append("</p>");
            } else if (ibodyelement.getElementType().equals(BodyElementType.TABLE)) {
                XWPFTable table = (XWPFTable) ibodyelement;
                allHTML.append("<table border=1>");
                for (XWPFTableRow row : table.getRows()) {
                    allHTML.append("<tr>");
                    for (XWPFTableCell cell : row.getTableCells()) {
                        allHTML.append("<td>");
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            allHTML.append("<p>");
                            allHTML.append(getTextAndFormulas(paragraph));
                            allHTML.append("</p>");
                        }
                        allHTML.append("</td>");
                    }
                    allHTML.append("</tr>");
                }
                allHTML.append("</table>");
            }
        }

        String[] data = allHTML.toString().split("#");
        List<String> keys = new ArrayList<>();
        List<String> values = new ArrayList<>();
        for (int i = 0; i < data.length - 1; i++) {
            String[] datss = data[i].split("\n");
            String[] answers = datss[0].split("\\.");
            if (answers.length == 2) {
                keys.add(answers[1]);
            }
            StringBuilder questions = new StringBuilder();

            for (int i1 = 1; i1 < datss.length; i1++) {
                questions.append(datss[i1]);
            }

            if (questions.length() > 1) {
                values.add(questions.toString());

            }
        }

        for (String value : values) {
            System.out.println("paragraph -> " + value);
        }

        document.close();

        //creating a sample HTML file
        String encoding = "UTF-8";
        FileOutputStream fos = new FileOutputStream("C:\\Users\\User\\Desktop\\result.html");
        OutputStreamWriter writer = new OutputStreamWriter(fos, encoding);
        writer.write("<!DOCTYPE html>\n");
        writer.write("<html lang=\"en\">");
        writer.write("<head>");
        writer.write("<meta charset=\"utf-8\"/>");

        //using MathJax for helping all browsers to interpret MathML
        writer.write("<script type=\"text/javascript\"");
        writer.write(" async src=\"https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.1/MathJax.js?config=MML_CHTML\"");
        writer.write(">");
        writer.write("</script>");

        writer.write("</head>");
        writer.write("<body>");

        writer.write(allHTML.toString());

        writer.write("</body>");
        writer.write("</html>");
        writer.close();

        Desktop.getDesktop().browse(new File("C:\\Users\\User\\Desktop\\result.html").toURI());

    }
}
