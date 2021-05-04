package excelCreating;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriteFormulaCell {
    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Numbers");

        XSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue(100);
        row.createCell(1).setCellValue(200);
        row.createCell(2).setCellValue(300);

        row.createCell(3).setCellFormula("A1*B1*C1");


        String filePath = ".\\datafiles\\Calc.xlsx";
        FileOutputStream fos = new FileOutputStream(filePath);

        workbook.write(fos);

        fos.close();

        System.out.println("Calc.xlsx file written is successful");



    }
}
