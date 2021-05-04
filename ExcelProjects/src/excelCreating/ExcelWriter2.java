package excelCreating;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class ExcelWriter2 {
    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Emp info");

        ArrayList<Object[]> empdata = new ArrayList<>();
        empdata.add(new Object[]{"Emp Id", "Name", "Job"});
        empdata.add(new Object[]{101, "Maciek", "Ogrodnik"});
        empdata.add(new Object[]{102, "Magda", "Korpolidek"});
        empdata.add(new Object[]{103, "Jacek", "Policjant"});


        //use for...each loop

        int rownum = 0;

        for (Object[] emp:empdata) {
            XSSFRow row = sheet.createRow(rownum++);
            int cellnum = 0;
            for (Object value:emp) {
                XSSFCell cell = row.createCell(cellnum++);
                if(value instanceof String) cell.setCellValue((String) value);
                if(value instanceof Integer) cell.setCellValue((Integer) value);
                if(value instanceof Boolean) cell.setCellValue((Boolean) value);

            }
        }
        String filePath = ".\\datafiles\\employee4.xlsx";
        FileOutputStream outputStream = new FileOutputStream(filePath);
        workbook.write(outputStream);

        outputStream.close();

        System.out.println("Employee.xlsx file written is successful");

    }
}
