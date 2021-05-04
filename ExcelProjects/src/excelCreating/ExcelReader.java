package excelCreating;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ExcelReader {
    public static void main(String[] args) throws IOException {
        String excelFilePath = ".\\datafiles\\countries.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);

        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

        XSSFSheet sheet = workbook.getSheetAt(0);
        //XSSFSheet sheet = workbook.getSheet("Sheet1");

        //using for loop
        /*int rows = sheet.getLastRowNum();
        int cols = sheet.getRow(1).getLastCellNum();

        for (int i = 0; i <= rows; i++) {
            XSSFRow row = sheet.getRow(i);
            for (int j = 0; j < cols ; j++) {
               XSSFCell cell = row.getCell(j);

               switch (cell.getCellType())
               {
                   case STRING:
                       System.out.print(cell.getStringCellValue()); break;
                   case NUMERIC:
                       System.out.print(cell.getNumericCellValue()); break;
                   case BOOLEAN:
                       System.out.print(cell.getBooleanCellValue()); break;
               }
                System.out.print(" | ");
            }
            System.out.println();
        }*/

        //iterator

       Iterator iterator = sheet.iterator();

        while (iterator.hasNext()){
            XSSFRow row = (XSSFRow) iterator.next();
            Iterator cellIterator =row.cellIterator();

            while (cellIterator.hasNext()){
                XSSFCell cell = (XSSFCell) cellIterator.next();
                switch (cell.getCellType())
                {
                    case STRING:
                        System.out.print(cell.getStringCellValue()); break;
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue()); break;
                    case BOOLEAN:
                        System.out.print(cell.getBooleanCellValue()); break;
                }
                System.out.print(" | ");
            }
            System.out.println();

        }


    }

}
