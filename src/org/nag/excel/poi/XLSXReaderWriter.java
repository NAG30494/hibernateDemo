package org.nag.excel.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSXReaderWriter {
	public static void main(String[] args) {
		try {
			File excel = new File("E://temp/Employee.xlsx");
			FileInputStream fis = new FileInputStream(excel);
			XSSFWorkbook book = new XSSFWorkbook(fis);
			XSSFSheet sheet = book.getSheetAt(0);
			Iterator<Row> itr = sheet.iterator(); // Iterating over Excel file in Java
			while (itr.hasNext()) {
				Row row = itr.next(); // Iterating over each column of Excel file
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue() + "\t");
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t");
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						System.out.print(cell.getBooleanCellValue() + "\t");
						break;
					default:
					}
				}
				System.out.println("");
			} // writing data into XLSX file
			Map<String, Object[]> newData = new HashMap<String, Object[]>();
			newData.put("7", new Object[] { 7d, "Sonya", "75K", "SALES","Rupert" });
			newData.put("8", new Object[] { 8d, "Kris", "85K", "SALES",	"Rupert" });
			newData.put("9", new Object[] { 9d, "Dave", "90K", "SALES",	"Rupert" });
			Set<String> newRows = newData.keySet();
			int rownum = sheet.getLastRowNum();
			for (String key : newRows) {
				Row row = sheet.createRow(rownum++);
				Object[] objArr = newData.get(key);
				int cellnum = 0;
				for (Object obj : objArr) {
					Cell cell = row.createCell(cellnum++);
					if (obj instanceof String) {
						cell.setCellValue((String) obj);
					} else if (obj instanceof Boolean) {
						cell.setCellValue((Boolean) obj);
					} else if (obj instanceof Date) {
						cell.setCellValue((Date) obj);
					} else if (obj instanceof Double) {
						cell.setCellValue((Double) obj);
					}
				}
			} // open an OutputStream to save written data into Excel file
			FileOutputStream os = new FileOutputStream(excel);
			book.write(os);
			System.out.println("Writing on Excel file Finished ..."); 
			book.close();
			fis.close();
		} catch (FileNotFoundException fe) {
			fe.printStackTrace();
		} catch (IOException ie) {
			ie.printStackTrace();
		}
	}
}
