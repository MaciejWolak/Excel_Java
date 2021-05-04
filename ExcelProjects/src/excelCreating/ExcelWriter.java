package excelCreating;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriter {
    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Emp info");

        Object[][] empdata = {{"Emp Id", "Name", "Job", }, {101,"David","Engenier"},
                {102,"Sam","Doctor"},{103,"George","Mechanic"}
        };

        //using for loop

       /* int rows = empdata.length; //4
        int cols = empdata[0].length; //3

        System.out.println(rows);
        System.out.println(cols);

        for (int i = 0; i < rows ; i++) {
            XSSFRow row = sheet.createRow(i);

            for (int j = 0; j < cols; j++) {
                XSSFCell cell = row.createCell(j);
                Object value = empdata[i][j];
                if(value instanceof String) cell.setCellValue((String) value);
                if(value instanceof Integer) cell.setCellValue((Integer) value);
                if(value instanceof Boolean) cell.setCellValue((Boolean) value);
            }
        }*/

        //use for...each loop

        int rowCount = 0;
        for (Object[] emp :empdata) {
            XSSFRow row = sheet.createRow(rowCount++);
            int columnCount = 0;
            for (Object value:emp) {
               XSSFCell cell = row.createCell(columnCount++);
                if(value instanceof String) cell.setCellValue((String) value);
                if(value instanceof Integer) cell.setCellValue((Integer) value);
                if(value instanceof Boolean) cell.setCellValue((Boolean) value);

            }
        }
        

        String filePath = ".\\datafiles\\employee2.xlsx";
        FileOutputStream outputStream = new FileOutputStream(filePath);
        workbook.write(outputStream);

        outputStream.close();

        System.out.println("Employee.xlsx file written is successful");

    }
}
