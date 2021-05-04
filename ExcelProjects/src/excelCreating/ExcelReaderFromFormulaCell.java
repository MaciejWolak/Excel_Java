package excelCreating;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReaderFromFormulaCell {
    public static void main(String[] args) throws IOException {

        FileInputStream file = new FileInputStream(".\\datafiles\\readformula.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        int rows =sheet.getLastRowNum();
        int colls = sheet.getRow(0).getLastCellNum();

        for (int i = 0; i <=rows ; i++) {
            XSSFRow row = sheet.getRow(i);
            for (int j = 0; j < colls; j++) {
                XSSFCell cell = row.getCell(j);
                switch (cell.getCellType())
                {
                    case STRING:
                        System.out.print(cell.getStringCellValue()); break;
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue()); break;
                    case BOOLEAN:
                        System.out.print(cell.getBooleanCellValue()); break;
                    case FORMULA:
                        System.out.println(cell.getNumericCellValue());
                }
                System.out.print(" | ");

            }
            System.out.println();
        }


    }
}
