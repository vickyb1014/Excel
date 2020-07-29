import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class AmazonProject {
static WebDriver driver;
	
@BeforeClass
private void browserLaunch() {
System.setProperty("webdriver.chrome.driver", "C:\\Vicky\\July2k20\\Test\\Drivers\\chromedriver.exe");
driver = new ChromeDriver();
driver.get("https://www.amazon.in/");
}
@Test
private void test() throws Throwable {
File f = new File("C:\\Vicky\\July2k20\\Test\\Files\\Excel.xlsx");
FileInputStream fis = new FileInputStream(f);
XSSFWorkbook w = new XSSFWorkbook(fis);
Sheet s = w.getSheet("Sheet1");
for(int i=0; i<s.getPhysicalNumberOfRows();i++) {
Row r = s.getRow(i);
for(int j=0; j<r.getPhysicalNumberOfCells();j++) {
Cell c = r.getCell(j);
int ct = c.getCellType();
if(ct==1) {
	String cv = c.getStringCellValue();
	System.out.println(cv);
}
else if(ct==0) {
if(DateUtil.isCellDateFormatted(c)) {
Date d = c.getDateCellValue();
System.out.println(d);
}
else {
double cv = c.getNumericCellValue();
Long l = (long)cv;
String string = String.valueOf(l);
System.out.println(string);	
}
}
}
}
}

@AfterClass
private void quitBrowser() throws Throwable {
Thread.sleep(3000);
driver.quit();
}	
}
