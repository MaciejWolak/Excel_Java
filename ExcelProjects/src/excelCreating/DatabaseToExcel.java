package excelCreating;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;

public class DatabaseToExcel {
    public static void main(String[] args) throws SQLException, IOException {

        //connect to database
        Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/world?serverTimezone=UTC","root","");

        //statement/query
        Statement stmt = con.createStatement();
        ResultSet rs = stmt.executeQuery("select * from locations");

        //Excel
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Locations data");

        XSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("LOCATION_ID");
        row.createCell(1).setCellValue("STREET_ADDRESS");
        row.createCell(2).setCellValue("POSTAL_CODE");
        row.createCell(3).setCellValue("CITY");
        row.createCell(4).setCellValue("STATE_PROVINCE");
        row.createCell(5).setCellValue("COUNTRY_ID");

        int r = 1;
        while(rs.next()){
            double locId = rs.getDouble("LOCATION_ID");
            String streetAddress = rs.getNString("STREET_ADDRESS");
            String postalCode = rs.getNString("POSTAL_CODE");
            String city = rs.getNString("CITY");
            String stateProvince = rs.getNString("STATE_PROVINCE");
            String countryId = rs.getNString("COUNTRY_ID");

            row = sheet.createRow(r++);

            row.createCell(0).setCellValue(locId);
            row.createCell(1).setCellValue(streetAddress);
            row.createCell(2).setCellValue(postalCode);
            row.createCell(3).setCellValue(city);
            row.createCell(4).setCellValue(stateProvince);
            row.createCell(5).setCellValue(countryId);


        }

        FileOutputStream fos = new FileOutputStream(".\\datafiles\\location3.xlsx");
        workbook.write(fos);

        workbook.close();
        fos.close();
        con.close();

        System.out.println("Done!!");






    }
}
