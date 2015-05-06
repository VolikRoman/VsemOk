package functional;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelTest {
    @Test
    public void test() throws IOException {
        WebDriver driver = new FirefoxDriver();

        driver.get("http://en.wikipedia.org/wiki/Software_testing");
        String title = driver.getTitle();
        System.out.println(title);

        FileInputStream input = new FileInputStream("D:\\sel.xls");

        int count=0;

        HSSFWorkbook wb = new HSSFWorkbook(input);
        HSSFSheet sh = wb.getSheet("sheet1");
        HSSFRow row = sh.getRow(count);
        String data = row.getCell(1).toString();
        System.out.println(data);

        FileOutputStream webdata = new FileOutputStream ("D:\\sel.xls");
        if(title.equals(data)) {
            row.createCell(10).setCellValue("TRUE");
            wb.write(webdata);
        }
        else {
            row.createCell(11).setCellValue("FALSE");
            wb.write(webdata);
        }

        driver.close();
        wb.close();
        input.close();
    }

}
