import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

// Press Shift twice to open the Search Everywhere dialog and type `show whitespaces`,
// then press Enter. You can now see whitespace characters in your code.
public class Main {
    public static void main(String[] args) {
        // Press Alt+Enter with your caret at the highlighted text to see how
        // IntelliJ IDEA suggests fixing it.
       // System.out.printf("Hello and welcome!");

        // Press Shift+F10 or click the green arrow button in the gutter to run the code.
        //for (int i = 1; i <= 5; i++) {

            // Press Shift+F9 to start debugging your code. We have set one breakpoint
            // for you, but you can always add more by pressing Ctrl+F8.
            //System.out.println("i = " + i);
        String path = "d:\\erdr\\list.xlsx";
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(path);
        }catch (FileNotFoundException fnfe){
            System.out.println(fnfe);
        }

        XSSFWorkbook xssfWorkbook = null;
        try {
            xssfWorkbook = new XSSFWorkbook(fileInputStream);
        } catch (IOException ioe){
            System.out.println(ioe);
        }

//        Sheet sheet = (Sheet) xssfWorkbook.getSheetAt(0);
//        Row row = sheet.ge
        org.apache.poi.ss.usermodel.Sheet sheet = xssfWorkbook.getSheetAt(0);
//        org.apache.poi.ss.usermodel.Row row = sheet.getRow(0);
        Iterator rowIter = sheet.rowIterator();
        while (rowIter.hasNext()){
            org.apache.poi.ss.usermodel.Row row = (org.apache.poi.ss.usermodel.Row) rowIter.next();
            String cellText = "";
            org.apache.poi.ss.usermodel.Cell cell = row.getCell(0);
            cellText = cell.getStringCellValue();
            System.out.println(cellText);

        }
    }
}