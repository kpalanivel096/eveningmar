package org.maven;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Parent {
public static	WebDriver driver;
	public static void chromebrowser(String url) {
		WebDriverManager.chromedriver().setup();
		 driver = new ChromeDriver();
	driver.manage().window().maximize();
	driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
	driver.get(url);
	}
public static void textpass(WebElement send,String id) {
	send.sendKeys(id);
}
	
public static void valuepass(WebElement send,String pass1)	{
	send.sendKeys(pass1);
}
public static String excelData(String sheetName, int rowNo,int cellNo) throws IOException  {
	
File file=new File("C:\\Users\\palanivel\\eclipse-workspace\\Sample\\src\\test\\resources\\file2.xlsx");

FileInputStream fis = new FileInputStream(file);
Workbook book = new XSSFWorkbook(fis);
Sheet sheet = book.getSheet(sheetName);
Row row = sheet.getRow(rowNo);
Cell cell = row.getCell(cellNo);
int type = cell.getCellType();
String value ="";
if (type==1) {
value=cell.getStringCellValue();
}else if (DateUtil.isCellDateFormatted(cell)) {
Date date = cell.getDateCellValue();
SimpleDateFormat s = new SimpleDateFormat("dd,mmmm,yyyy");
value = s.format(date);
}else {
	double d =cell.getNumericCellValue();
	long l =(long) d;
value = String.valueOf(l);	

}
return value;
}

}
	
	
	
	
	
