package GPT_bigReader;


import org.apache.logging.log4j.core.util.JsonUtils;
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
//enp 01.08.2025 {
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
//enp}

public class ExcelMultiColumnToCsv {

    public static void main(String[] args) throws Exception {
        System.setOut(new java.io.PrintStream(System.out, true, "UTF-8"));
        String excelPath = "c:\\erdr\\erdr.xlsx";
        String csvOutputPath = "c:\\erdr\\output.csv";
        //enp{
        String listPath = "c:\\erdr\\list.xlsx";
        int searchColumnIndex1 = 0; //в файле list.xlsx всего одна колонка со списком ФИО
        //Формирование списка искомых военнослужащих
        List<String> stringsToFind = readColumn(listPath, searchColumnIndex1, 0);
        //enp}

        //enp 2026-07-18 у файлі list.xlsx додана вкладка NotHandlingNumbers
        //це список номерів ЄРДР що не треба обробляти, оскільки це повні тазки наших військовослужбовців
        //та не належать до нашої частини
        //сформуємо список номерів ЄРДР для пропуску при обробці
        searchColumnIndex1 = 0;
        List<String> stringErdrNumbersToSkip = readColumn(listPath, searchColumnIndex1, 1);
        System.out.println(stringErdrNumbersToSkip);

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

        List<String> militaryPersonnelFound;
        militaryPersonnelFound = new ArrayList<>();
        XMLReader parser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
        parser.setContentHandler(new MultiColumnHandler(sst, targetColumnIndices, resultRows, stringsToFind,stringErdrNumbersToSkip, militaryPersonnelFound));
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
//        //enp{
//        // Запись в CSV с кодировкой UTF-8
//        FileOutputStream fos = new FileOutputStream(csvOutputPath);
//        fos.write(0xEF); fos.write(0xBB); fos.write(0xBF); // UTF-8 BOM
//        try (Writer writer = new OutputStreamWriter(fos, StandardCharsets.UTF_8)) {
//            for (List<String> row : resultRows) {
//                writer.write(String.join("\t", row).replace("\"", "\"\""));
//                writer.write("\n");
//            }
//            System.out.println("✅ CSV-файл сохранен в UTF-8: " + csvOutputPath);
//        }
//        //enp}


//        //enp 31072025 {
//        String xlsxOutputPath = "c:\\erdr\\output.xlsx";
//        Workbook workbookOut = new XSSFWorkbook();
//        Sheet sheetOut = workbookOut.createSheet("Results");
//
//        for (int i = 0; i < resultRows.size(); i++) {
//            Row row = sheetOut.createRow(i);
//            List<String> dataRow = resultRows.get(i);
//            for (int j = 0; j < dataRow.size(); j++) {
//                Cell cell = row.createCell(j);
//                cell.setCellValue(dataRow.get(j));
//            }
//        }
//
//        try (FileOutputStream fileOut = new FileOutputStream(xlsxOutputPath)) {
//            workbookOut.write(fileOut);
//            workbookOut.close();
//            System.out.println("✅ XLSX-файл збережено: " + xlsxOutputPath);
//        }
//        //enp}


        //enp 01.08.2025{
        String xlsxOutputPath = "c:\\erdr\\output.xlsx";
        Workbook workbookOut = new XSSFWorkbook();
        Sheet sheetOut = workbookOut.createSheet("Results");

        // Створюємо стиль дати
        CreationHelper createHelper = workbookOut.getCreationHelper();
        CellStyle dateCellStyle = workbookOut.createCellStyle();
        short dateFormat = createHelper.createDataFormat().getFormat("dd.MM.yyyy");
        dateCellStyle.setDataFormat(dateFormat);

        for (int i = 0; i < resultRows.size(); i++) {
            Row row = sheetOut.createRow(i);
            List<String> dataRow = resultRows.get(i);
            for (int j = 0; j < dataRow.size(); j++) {
                Cell cell = row.createCell(j);
                String cellValue = dataRow.get(j);

                // 🔍 Спроба розпізнати дату
                if (cellValue.matches("\\d{2}\\-\\d{2}\\-\\d{4} \\d{2}:\\d{2}:\\d{2}") || cellValue.matches("\\d{2}\\.\\d{2}\\.\\d{4} \\d{2}:\\d{2}:\\d{2}")) {
//                    System.out.println("Знайдено комірку з форматом дати:");
//                    System.out.println(cellValue);
                    try {
//                        java.util.Date date = java.sql.Date.valueOf(cellValue.split(" ")[0].replace(".", "-"));
//                        cell.setCellValue(date);
//                        cell.setCellStyle(dateCellStyle);
                        //enp 01082025 21:29 {
                       //String fullDateTime = "23.07.2025 11:40:05";
                        String fullDateTime = cellValue;

                        // Формат вхідної дати
                        SimpleDateFormat fullFormat = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss");

                        // Формат вихідної дати
                        SimpleDateFormat shortFormat = new SimpleDateFormat("dd.MM.yyyy");

                        try {
                            // Парсимо повну дату у об'єкт Date
                            Date date_ = fullFormat.parse(fullDateTime);

                            // Форматуємо у короткий вигляд
                            String shortDate = shortFormat.format(date_);

                            //System.out.println("👉 Коротка дата: " + shortDate);
                            cell.setCellValue(shortDate);
                        } catch (ParseException e) {
                            System.out.println("❌ Помилка при обробці дати: " + e.getMessage());
                        }
                        //enp }
                    } catch (Exception e) {
                        // Якщо не вдалося — зберегти як текст
                        cell.setCellValue(cellValue);
                    }
                } else {
                    cell.setCellValue(cellValue);
                }
            }
        }

        //enp 03082025{
        Sheet sheetFoundPersonnel = workbookOut.createSheet("FoundPersonnel");
        for (int i = 0; i < militaryPersonnelFound.size(); i++) {
            Row row = sheetFoundPersonnel.createRow(i);
            Cell cell = row.createCell(0);
            String cellValue = militaryPersonnelFound.get(i);
            cell.setCellValue(cellValue);
        }
        //enp}

        try (FileOutputStream fileOut = new FileOutputStream(xlsxOutputPath)) {
            workbookOut.write(fileOut);
            workbookOut.close();
            System.out.println("✅ XLSX-файл збережено з форматуванням дат: " + xlsxOutputPath);
        }
        //enp}

//        //enp 03082025 {
//        System.out.println("List militaryPersonnelFound:");
//        System.out.println(militaryPersonnelFound);
//        //enp}
    }

    public static List<String> readColumn(String filePath, int columnIndex, int sheetNumber) {
        //enp {добавил метод для формирования списка искомых военнослужащих}
        List<String> columnData = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(sheetNumber); // Читаємо вказаний лист

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
        private final List<String> listToSkip;
        //enp}
        //enp 03082025 {
        private final List<String> militaryPersonnelFound;
        //enp}
        private final Map<Integer, String> currentRow = new TreeMap<>();

        private String currentCellRef;
        private boolean isString;
        private int currentColumn = -1;
        private StringBuilder value = new StringBuilder();

        MultiColumnHandler(SharedStrings sst, Set<Integer> targetColumns, List<List<String>> resultRows, List<String> listToFind, List<String> listToSkip, List<String> militaryPersonnelFound) {
            this.sst = sst;
            this.targetColumns = targetColumns;
            this.resultRows = resultRows;
            this.listToFind = listToFind;
            this.listToSkip = listToSkip;
            this.militaryPersonnelFound = militaryPersonnelFound;
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
                // Колонку фабула для гнучкості пошуку приводимо у нижній регістр
                // Прочитаємо окремі колонки з обраного рядка
                int columnNumberToSearch = 10;
                int columnNumberErdr = 7;
                String fabulaToSearch = rowOutput.get(columnNumberToSearch);
                String erdr = rowOutput.get(columnNumberErdr);
                for (String militaryPersonnelToSearch:listToFind){
                    if(fabulaToSearch.toLowerCase().contains(militaryPersonnelToSearch.toLowerCase())) {
                        //enp 2026-07-18 додамо перевірку для пропуску військвослужбовців, що не належать до нашої частини
                        //{
                        boolean skipErdr = false;
                        for (String erdrNumbersToSkip:listToSkip) {
                            if (erdr.toLowerCase().contains(erdrNumbersToSkip.toLowerCase())){
                                skipErdr = true;
                            }
                        }
                        if(skipErdr){
                            System.out.println("This row will be skipped -----");
                            System.out.println(rowOutput);
                            System.out.println("----");
                        }else {

                            //}
                            resultRows.add(rowOutput);
                            System.out.println(militaryPersonnelToSearch);
                            militaryPersonnelFound.add(militaryPersonnelToSearch);
                        }
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