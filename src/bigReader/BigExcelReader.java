package bigReader;

import org.apache.poi.openxml4j.opc.OPCPackage;
        import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
        import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.xml.sax.Attributes;
        import org.xml.sax.InputSource;
        import org.xml.sax.XMLReader;
        import org.xml.sax.helpers.DefaultHandler;

        import javax.xml.parsers.SAXParserFactory;
        import java.io.InputStream;

public class BigExcelReader {
    public static void main(String[] args) throws Exception {
        String path = "d:\\erdr\\erdrToExcel\\erdr.xlsx";
        OPCPackage pkg = OPCPackage.open(path);
        XSSFReader reader = new XSSFReader(pkg);
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);

        InputStream sheet = reader.getSheetsData().next();
        InputSource source = new InputSource(sheet);

        XMLReader parser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();

        parser.setContentHandler(new DefaultHandler() {
            public void startElement(String uri, String localName, String qName, Attributes attributes) {
                if ("row".equals(qName)) {
                    System.out.println("🔹 Початок рядка");
                }
            }

            public void characters(char[] ch, int start, int length) {
                String data = new String(ch, start, length).trim();
                if (!data.isEmpty()) {
                    System.out.println("   ▪️ Значення: " + data);
                }
            }

            public void endElement(String uri, String localName, String qName) {
                if ("row".equals(qName)) {
                    System.out.println("🔻 Кінець рядка\n");
                }
            }
        });

        parser.parse(source);
        sheet.close();
    }
}
