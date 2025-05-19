package XSSF_reader;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Working {




    public static void main(String[] args) {
        String filePath1 = "d:\\erdr\\list.xlsx "; // Замените на фактический путь к первому файлу
        String filePath2 = "d:\\erdr\\erdr.xlsx "; // Замените на фактический путь ко второму файлу
        String filePathRes = "d:\\erdr\\res.xlsx "; // strings that were founded in erdr

        //{enp
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        // write to file
        try (FileOutputStream fileOut = new FileOutputStream(filePathRes)){
            workbook.write(fileOut);
            workbook.close();

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        //enp}

        int searchColumnIndex1 = 0; // Индекс столбца для поиска в первом файле (начинается с 0)
        int searchColumnIndex2 = 0; // Индекс столбца для поиска во втором файле (начинается с 0) !! фабула 10

        List<String> stringsToFind = readColumn(filePath1, searchColumnIndex1);
        List<String> foundStrings = findStrings(filePath2, searchColumnIndex2, stringsToFind);

        if (foundStrings.isEmpty()) {
            System.out.println("Совпадения не найдены.");
        } else {
            System.out.println("Найденные строки во втором файле:");
            for (String foundString : foundStrings) {
                System.out.println(foundString);
            }
        }
    }

    public static List<String> readColumn(String filePath, int columnIndex) {
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

    public static List<String> findStrings(String filePath, int columnIndex, List<String> searchList) {
        List<String> foundStrings = new ArrayList<>();
        //{enp
        IOUtils.setByteArrayMaxOverride(250 * 1024 * 1024); // дозволити до 250 МБ
        //enp}
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook (fis)) {
            Sheet sheet = workbook.getSheetAt(0); // Читаем первый лист

            for (Row row : sheet) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    String cellValue = "";
                    switch (cell.getCellType()) {
                        case STRING:
                            cellValue = cell.getStringCellValue().trim();
                            break;
                        case NUMERIC:
                            cellValue = String.valueOf(cell.getNumericCellValue()).trim();
                            break;
                        case BOOLEAN:
                            cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
                            break;
                        case FORMULA:
                            try {
                                cellValue = cell.getStringCellValue().trim();
                            } catch (IllegalStateException e) {
                                cellValue = String.valueOf(cell.getNumericCellValue()).trim();
                            }
                            break;
                        default:
                            continue;
                    }
//                        if (searchList.contains(cellValue)) {
//                            foundStrings.add(cellValue);
//                        }
                    //{enp
                    for (String element:searchList){
                        if (cellValue.toLowerCase().contains(element.toLowerCase())) {
                            foundStrings.add(element);
                        }
                    }
                    //enp}
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return foundStrings;
    }


}