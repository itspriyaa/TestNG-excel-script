package automationFramework;
import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class login1 {
WebDriver driver;
@BeforeMethod
public void beforeMethod() throws Exception {
String path="D:\\Selenium\\chromedriver_win32\\chromedriver.exe";
System.setProperty("webdriver.chrome.driver", path);
driver=new ChromeDriver();
driver.get("file:///C:/Users/itctesting02/Downloads/Portfolio-Web_Application-Project-main/Portfolio-WebApplication/Registration%20Form/Login%20Form/index.html?");
}
@Test(dataProvider="getData")
public void test(String email,String password) {
driver.findElement(By.id("phone-email")).sendKeys(email);
driver.findElement(By.id("password")).sendKeys(password);
driver.findElement(By.id("submittion")).click();

}
@DataProvider
public String[][] getData() throws Exception {
File src=new File("C:\\Users\\itctesting02\\Desktop\\Test case and Test Data\\log1.xlsx");
FileInputStream fis=new FileInputStream(src);
XSSFWorkbook wb=new XSSFWorkbook(fis);
XSSFSheet sheet=wb.getSheet("Sheet1");
int Rows=sheet.getPhysicalNumberOfRows();
int cols=sheet.getRow(0).getLastCellNum();

String[][] data=new String[Rows-1][cols];
for(int i=0;i<Rows-1;i++)
{
for(int j=0;j<cols;j++)
{
DataFormatter df=new DataFormatter();
data[i][j]= df.formatCellValue(sheet.getRow(i+1).getCell(j));

}
System.out.println();
}
wb.close();
fis.close();
return data;
}
@AfterMethod
public void afterMethod() {
driver.quit();
}
}

