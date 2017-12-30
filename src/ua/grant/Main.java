package ua.grant;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class Main {

    public static void main(String[] args) {
        String name;
        try {
            File file = new File(".\\names.xml");
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(file);
            NodeList nList = doc.getElementsByTagName("name");
            doc.getDocumentElement().normalize();
            for (int temp = 0; temp < nList.getLength(); temp++) {
                Node nNode = nList.item(temp);
                name = nNode.getTextContent();
                createDocx(name);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void createDocx(String name) {
        XWPFDocument doc = null;
        try {
            doc = new XWPFDocument(OPCPackage.open(".\\template.docx"));
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.contains("$name")) {
                        text = text.replace("$name", name);
                        r.setText(text, 0);
                    }
                }
            }
        }

        try {
            doc.write(new FileOutputStream(".\\name_" + name + ".docx"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
