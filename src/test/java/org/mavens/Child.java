package org.mavens;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Child {
	private static final String FILE_PATH= "\\C:\\Users\\palanivel\\eclipse-workspace\\Sample\\src\\test\\resources\\file2.xlsx";

public static void main(String[] args) throws IOException {
/*	chromebrowser("https://www.fashionnova.com/");
	WebElement mailid = driver.findElement(By.id("email"));
File fi=new File("C:\\Users\\palanivel\\eclipse-workspace\\Sample\\src\\test\\resources\\file1.xlsx");
FileInputStream fis=new FileInputStream(fi);
Workbook book = new XSSFWorkbook(fis);
Sheet sheet = book.getSheet("sheet1");
Row row = sheet.getRow(1);
String id = row.getCell(1).getStringCellValue();
textpass(mailid, id);
row.getCell(2);*/
	


File f = new File(FILE_PATH);

FileInputStream fiss = new FileInputStream(f);
Workbook workbook = new XSSFWorkbook(fiss);


Sheet sheet = workbook.getSheet("sheet1");
Row createRow = sheet.createRow(11);
Cell c4 = createRow.createCell(3);
Cell c5 = createRow.createCell(4);
Cell c6 = createRow.createCell(5);
Cell c7 = createRow.createCell(6);
Cell c8 = createRow.createCell(7);

c4.setCellValue("kamalesh");
c5.setCellValue("prathivi");
c6.setCellValue("kamal");
c7.setCellValue("bharathi");
c8.setCellValue("banu");

FileOutputStream fo = new FileOutputStream(f);
workbook.write(fo);
int physicalNumberofRows = sheet.getPhysicalNumberOfRows();
System.out.println("Total Numbeer of Rows="+physicalNumberofRows);
int totalCells = 0;
for (int i = 0; i < physicalNumberofRows; i++) {
	Row row = sheet.getRow(i);
if (row!=null) {
	totalCells += row.getLastCellNum();
	
}
	
}
System.out.println("Total Number of Cells]="+totalCells);
/*for (int k = 0; k <physicalNumberofRows ; k++) {
	Row row = sheet.getRow(k);
}*/
}

}







	
