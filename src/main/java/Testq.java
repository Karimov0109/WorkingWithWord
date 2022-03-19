////import org.apache.poi.xwpf.usermodel.XWPFDocument;
////import org.apache.poi.xwpf.usermodel.XWPFParagraph;
////import org.apache.poi.xwpf.usermodel.XWPFRun;
////
////import java.io.File;
////import java.io.FileInputStream;
////import java.io.FileOutputStream;
////import java.io.IOException;
////import java.util.LinkedList;
////import java.util.List;
////
////public class Testq {
////    public static void main(String[] args) {
////        //exampleFile is the layout file you provided with data added for testing
////        List<String> values;
////        values = parseWordDocument("C:\\Users\\User\\Desktop\\Matematika_baza.docx");
////
////        for(String s: values)
////            System.out.println(s);
////    }
////
////    public static List<String> parseWordDocument(String documentPath) {
////        FileInputStream fInput = null;
////        XWPFDocument document = null;
////        List<String> parsedValues = null;
////
////        try {
////            File file = new File(documentPath);
////
////            fInput = new FileInputStream(file.getAbsolutePath());
////            document = new XWPFDocument(fInput);
////
////            //getParagraphs() will grab each paragraph for you
////            List<XWPFParagraph> paragraphs = document.getParagraphs();
////
////            parsedValues = new LinkedList<>();
////
////            for (XWPFParagraph para : paragraphs) {
////                //remove the title
////                if(!para.getText().startsWith("#")) {
////                    //here is where you want to parse your line to get needed values
////                    String[] splitLine = para.getText().split("#");
////                    //based on example input file [1] is the value you need
////                    parsedValues.add(splitLine[0]);
////                }
////            }
////            System.out.println(parsedValues.get(0));
////
////            fInput.close();
////            document.close();
////        } catch (Exception e) {
////            e.printStackTrace();
////        }
////        return parsedValues;
////    }
////}
//for (int i = 0; i < 5; i++) {
//        int random_int1 = (int) Math.floor(Math.random() * (50) + (1));
//        System.out.println(values.get(random_int1));
//        }
//
//
//        int random_int2 = 0;
//        for (int i = 0; i < 5; i++) {
//        random_int2 = (int) Math.floor(Math.random() * (100) + (51));
//        System.out.println(values.get(random_int2));
//        }
//
//
//        for (int i = 0; i < 5; i++) {
//        int random_int3 = (int) Math.floor(Math.random() * (150) + (101));
//        System.out.println(values.get(random_int3));
//        }
//
//
//        for (int i = 0; i < 5; i++) {
//        int random_int4 = (int) Math.floor(Math.random() * (200) + (151));
//        System.out.println(values.get(random_int4));
//        }
//
//
//        for (int i = 0; i < 5; i++) {
//        int random_int5 = (int) Math.floor(Math.random() * (250) + (201));
//        System.out.println(values.get(random_int5));
//        }
//
//
//        for (int i = 0; i < 5; i++) {
//        int random_int6 = (int) Math.floor(Math.random() * (300) + (251));
//        System.out.println(values.get(random_int6));
//        }