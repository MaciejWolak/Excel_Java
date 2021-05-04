package excelCreating;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;

public class ExcelToDatabase {
    public static void main(String[] args) throws SQLException, IOException {

        //connect to database
        Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/world?serverTimezone=UTC", "root", "");

        Statement stmt = con.createStatement();

        //create new table in the databases

        String sql = "CREATE TABLE IF NOT EXISTS names(Id decimal(4,0), firstName varchar(255), lastName varchar(255), PRIMARY KEY(Id) ) ";
        stmt.execute(sql);

        //Excel

        FileInputStream fis = new FileInputStream(".\\datafiles\\ExcelToDatabase3.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        XSSFSheet sheet = workbook.getSheet("Arkusz1");

        int rows = sheet.getLastRowNum();

        for (int i = 1; i <= rows; i++) {
            XSSFRow row = sheet.getRow(i);
            double locId =row.getCell(0).getNumericCellValue();
            String firstname = row.getCell(1).getStringCellValue();
            String lastname = row.getCell(1).getStringCellValue();

            sql = "INSERT INTO names VALUES ('"+locId+"','"+firstname+"','"+lastname+"')";
            stmt.execute(sql);
            stmt.execute("commit");


        }
        workbook.close();
        fis.close();
        con.close();

        System.out.println("Done!");






    }
}