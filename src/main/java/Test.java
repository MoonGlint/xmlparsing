import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;

public class Test {
    private static Node node;
    private static Element element;
    private static NodeList subList;
    private static String textContent;
    private static ExcelDocCreator excelDocCreator;
    private static int k = 0;
    private static int count = 0;

    public static void main(String[] args) {
        /////////////////////
        excelDocCreator = new ExcelDocCreator();
        excelDocCreator.createDefaultSheet();
        /////////////////////


        Document document = null;
        DocumentBuilderFactory documentBuilderFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder documentBuilder = null;
        try {
            documentBuilder = documentBuilderFactory.newDocumentBuilder();
        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        }
        try {
            document = documentBuilder.parse("src\\main\\resources\\jun15march.xml");
            //optional, but recommended
            //read this - http://stackoverflow.com/questions/13786607/normalization-in-dom-parsing-with-java-how-does-it-work
            document.getDocumentElement().normalize();
        } catch (IOException | SAXException | NullPointerException e) {
            e.printStackTrace();
        }

        NodeList STATEMENTBYList = document.getElementsByTagName("STATEMENTBY");
        NodeList DEBETDOCUMENTSList = document.getElementsByTagName("DEBETDOCUMENTS");
        NodeList CREDITDOCUMENTSList = document.getElementsByTagName("CREDITDOCUMENTS");


        System.out.println(STATEMENTBYList.getLength());
        //int counter = 0;

        for (k = 0; k < STATEMENTBYList.getLength(); k++) {

            for (int j = 0; j < DEBETDOCUMENTSList.getLength(); j++) {
                node = DEBETDOCUMENTSList.item(k).getChildNodes().item(j);
                utilMeth(node, "DEBETDOCUMENTS");
            }

            for (int m = 0; m < CREDITDOCUMENTSList.getLength(); m++) {
                node = CREDITDOCUMENTSList.item(k).getChildNodes().item(m);
                utilMeth(node, "CREDITDOCUMENTS");
            }
            System.out.println("=====================================================================");
           /* for (int i = 0; i < STATEMENTBYList.item(k).getChildNodes().getLength(); i++) {

                node = STATEMENTBYList.item(k).getChildNodes().item(i);
                if ((node != null) && (node.getNodeType() == Node.ELEMENT_NODE)) {
                    element = (Element) node;
                    textContent = element.getTextContent();
                    switch (element.getNodeName()) {

                        case "STATEMENTDATE":
                            System.out.println("STATEMENTDATE " + element.getTextContent());
                            excelDocCreator.setCellData(textContent, k + 1, 6);
                            break;
                        case "OPENINGBALANCE":
                            System.out.println("OPENINGBALANCE " + element.getTextContent());
                            excelDocCreator.setCellData(textContent, k + 1, 7);
                            break;

                        case "CLOSINGBALANCE":
                            System.out.println("CLOSINGBALANCE " + element.getTextContent());
                            excelDocCreator.setCellData(textContent, k + 1, 10);
                            break;
                        case "DEBETTURNOVER":
                            System.out.println("DEBETTURNOVER " + element.getTextContent());
                            excelDocCreator.setCellData(textContent, k + 1, 8);
                            break;
                        case "CREDITTURNOVER":
                            System.out.println("CREDITTURNOVER " + element.getTextContent());
                            excelDocCreator.setCellData(textContent, k + 1, 9);
                            break;
                    }
                }
            }*/
            System.out.println("=====================================================================");
        }

        excelDocCreator.writeSheet();
    }

    private static void utilMeth(Node node, String currentDoc) {


        if ((node != null) && (node.getNodeType() == Node.ELEMENT_NODE)) {
            element = (Element) node;
            subList = node.getChildNodes();
            textContent = element.getTextContent();
            switch (currentDoc) {
                case ("DEBETDOCUMENTS"):
                    System.out.println();
                    System.out.println("DEBETDOCUMENTS");
                    System.out.println("----------------------------------");
                    break;
                case ("CREDITDOCUMENTS"):
                    System.out.println();
                    System.out.println("CREDITDOCUMENTS");
                    System.out.println("----------------------------------");
                    break;

            }

            for (int f = 0; f < subList.getLength(); f++) {
                node = subList.item(f);
                switch (node.getNodeName()) {
                    case "DOCUMENTNUMBER":
                        count++;
                        break;
                    case "DOCUMENTDATE":
                        System.out.println("DOCUMENTDATE " + node.getTextContent());

                        excelDocCreator.setCellData(node.getTextContent(), count, 0);
                        break;
                    case "PAYER":
                        System.out.println("PAYER " + node.getTextContent());

                        excelDocCreator.setCellData(node.getTextContent(), count, 4);
                        break;
                    case "AMOUNT":
                        System.out.println("AMOUNT " + node.getTextContent());
                        switch (currentDoc) {
                            case ("DEBETDOCUMENTS"):
                                try {
                                    excelDocCreator.setCellData(Double.parseDouble(node.getTextContent().replace(',', '.')), count, 1);

                                } catch (NumberFormatException nf) {
                                    nf.printStackTrace();
                                }
                                break;
                            case ("CREDITDOCUMENTS"):
                                try {
                                    excelDocCreator.setCellData(Double.parseDouble(node.getTextContent().replace(',', '.')), count, 2);
                                } catch (NumberFormatException nf) {
                                    nf.printStackTrace();
                                }
                                break;
                        }

                        break;
                    case "BENEFICIAR":
                        System.out.println("BENEFICIAR " + node.getTextContent());

                        excelDocCreator.setCellData(node.getTextContent(), count, 3);
                        break;
                    case "GROUND":
                        System.out.println("GROUND " + node.getTextContent());

                        excelDocCreator.setCellData(node.getTextContent(), count, 5);
                        break;
                }
            }
        }
    }
}




