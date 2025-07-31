package GPT_bigReader;


import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
        import org.apache.poi.xssf.eventusermodel.XSSFReader;
        import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.Attributes;
        import org.xml.sax.InputSource;
        import org.xml.sax.XMLReader;
        import org.xml.sax.helpers.DefaultHandler;

        import javax.xml.parsers.SAXParserFactory;
import java.io.*;
import java.util.*;
import java.io.Writer;
import java.nio.charset.StandardCharsets;

public class ExcelMultiColumnToCsv {

    public static void main(String[] args) throws Exception {
        String excelPath = "d:\\erdr\\erdr.xlsx";
        String csvOutputPath = "d:\\erdr\\output.csv";
        //enp{
        String listPath = "d:\\erdr\\list.xlsx";
        int searchColumnIndex1 = 0; //в файле list.xlsx всего одна колонка со списком ФИО
        //Формирование списка искомых военнослужащих
        List<String> stringsToFind = readColumn(listPath, searchColumnIndex1);
        //enp}

        // Указываем индексы нужных колонок (например: A, C, D)
        //enp{
        //Set<Integer> targetColumnIndices = new HashSet<>(Arrays.asList(0, 2, 3));
        Set<Integer> targetColumnIndices = new HashSet<>();

        for(int i = 0; i< 76; i++){
            targetColumnIndices.add(i);
        }
        //System.out.println(targetColumnIndices);
        //enp}

        List<List<String>> resultRows = new ArrayList<>();

        OPCPackage pkg = OPCPackage.open(excelPath);
        SharedStrings sst = new ReadOnlySharedStringsTable(pkg);
        XSSFReader xssfReader = new XSSFReader(pkg);
        InputStream sheet = xssfReader.getSheetsData().next();

        XMLReader parser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
        parser.setContentHandler(new MultiColumnHandler(sst, targetColumnIndices, resultRows, stringsToFind));
        parser.parse(new InputSource(sheet));

        sheet.close();
        pkg.close();

        /* // Запись в CSV
        try (FileWriter writer = new FileWriter(csvOutputPath)) {
            for (List<String> row : resultRows) {
                writer.write(String.join(";", row).replace("\"", "\"\""));
                writer.write("\n");
            }
            System.out.println("✅ CSV-файл сохранен: " + csvOutputPath);
        }

         */
        //enp{
        // Запись в CSV с кодировкой UTF-8
        FileOutputStream fos = new FileOutputStream(csvOutputPath);
        fos.write(0xEF); fos.write(0xBB); fos.write(0xBF); // UTF-8 BOM
        try (Writer writer = new OutputStreamWriter(fos, StandardCharsets.UTF_8)) {
            for (List<String> row : resultRows) {
                writer.write(String.join("\t", row).replace("\"", "\"\""));
                writer.write("\n");
            }
            System.out.println("✅ CSV-файл сохранен в UTF-8: " + csvOutputPath);
        }
        //enp}
    }

    public static List<String> readColumn(String filePath, int columnIndex) {
        //enp {добавил метод для формирования списка искомых военнослужащих}
        List<String> columnData = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0); // Читаем первый лист

            for (Row row : sheet) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    switch (cell.getCellType()) {
                        case STRING:
                            columnData.add(cell.getStringCellValue().trim());
                            break;
                        case NUMERIC:
                            columnData.add(String.valueOf(cell.getNumericCellValue()).trim());
                            break;
                        case BOOLEAN:
                            columnData.add(String.valueOf(cell.getBooleanCellValue()).trim());
                            break;
                        case FORMULA:
                            try {
                                columnData.add(cell.getStringCellValue().trim());
                            } catch (IllegalStateException e) {
                                columnData.add(String.valueOf(cell.getNumericCellValue()).trim());
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return columnData;
    }


    static class MultiColumnHandler extends DefaultHandler {
        private final SharedStrings sst;
        private final Set<Integer> targetColumns;
        private final List<List<String>> resultRows;
        //enp{
        private final List<String> listToFind;
        //enp}
        private final Map<Integer, String> currentRow = new TreeMap<>();

        private String currentCellRef;
        private boolean isString;
        private int currentColumn = -1;
        private StringBuilder value = new StringBuilder();

        MultiColumnHandler(SharedStrings sst, Set<Integer> targetColumns, List<List<String>> resultRows, List<String> listToFind) {
            this.sst = sst;
            this.targetColumns = targetColumns;
            this.resultRows = resultRows;
            this.listToFind = listToFind;
        }

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if ("row".equals(qName)) {
                currentRow.clear();
            }

            if ("c".equals(qName)) {
                currentCellRef = attributes.getValue("r"); // например, A1, C2
                String col = currentCellRef.replaceAll("[0-9]", "");
                currentColumn = getColumnIndex(col);
                isString = "s".equals(attributes.getValue("t"));
                value.setLength(0);
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            value.append(ch, start, length);
        }

        @Override
        public void endElement(String uri, String localName, String qName) {
            if ("v".equals(qName) && targetColumns.contains(currentColumn)) {
                String cellValue = value.toString().trim();
                if (isString) {
                    try {
                        int idx = Integer.parseInt(cellValue);
                        cellValue = sst.getItemAt(idx).getString();
                    } catch (NumberFormatException e) {
                        cellValue = "";
                    }
                }
                currentRow.put(currentColumn, cellValue);
            }

            if ("row".equals(qName)) {
                List<String> rowOutput = new ArrayList<>();
                for (int colIndex : targetColumns) {
                    rowOutput.add(currentRow.getOrDefault(colIndex, ""));
                }
                //enp{
                int columnNumberToSearch = 10;
                String fabulaToSearch = rowOutput.get(columnNumberToSearch);
                for (String militaryPersonnelToSearch:listToFind){
                    if(fabulaToSearch.toLowerCase().contains(militaryPersonnelToSearch.toLowerCase())) {
                        //System.out.println(rowOutput.get(columnNumberToSearch));
                        resultRows.add(rowOutput);
                        System.out.println(militaryPersonnelToSearch);
                    }
                }
                //enp}
                //resultRows.add(rowOutput);

            }
        }

        private int getColumnIndex(String column) {
            int index = 0;
            for (int i = 0; i < column.length(); i++) {
                index = index * 26 + (column.charAt(i) - 'A' + 1);
            }
            return index - 1;
        }
    }
}
//test