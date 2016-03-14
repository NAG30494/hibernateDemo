package org.nag.excel.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class WriteXLSX {
	public static void main(String[] args) throws IOException {
		FileInputStream fis = new FileInputStream(new File("e:\\temp\\income tax calculator for the financial year 2013-14.xlsx"));
		XSSFWorkbook workbook = new XSSFWorkbook (fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		//Create First Row
		//XSSFRow row1 = sheet.createRow(8);
		//XSSFCell cell = row1.createCell(4);
		XSSFRow row = sheet.getRow(8);
		XSSFCell cell = row.getCell(4);
		cell.setCellValue(250000.00);
		fis.close();
		FileOutputStream fos =new FileOutputStream(new File("e:\\temp\\income tax calculator for the financial year 2013-141.xlsx"));
	    workbook.write(fos);
	    fos.close();
		System.out.println("Done");
	}
} 