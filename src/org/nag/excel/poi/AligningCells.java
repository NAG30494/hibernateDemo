package org.nag.excel.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.impl.CTRowImpl;

public class AligningCells
{
  public static void main(String[] args)
    throws IOException
  {
    XSSFWorkbook wb = new XSSFWorkbook();

    XSSFSheet sheet = wb.createSheet();
    XSSFRow row = sheet.createRow(2);
    row.setHeightInPoints(30.0F);
    for (int i = 0; i < 8; i++)
    {
      sheet.setColumnWidth(i, 3840);
    }

    createCell(wb, row, (short)0, (short)2, (short)2);
    createCell(wb, row, (short)1, (short)6, (short)2);
    createCell(wb, row, (short)2, (short)4, (short)1);
    createCell(wb, row, (short)3, (short)0, (short)1);
    createCell(wb, row, (short)4, (short)5, (short)3);
    createCell(wb, row, (short)5, (short)1, (short)0);
    createCell(wb, row, (short)6, (short)3, (short)0);

    row = sheet.createRow(3);
    centerAcrossSelection(wb, row, (short)1, (short)3, (short)1);

    FileOutputStream fileOut = new FileOutputStream("xssf-align.xlsx");
    wb.write(fileOut);
    fileOut.close();
  }

  private static void createCell(XSSFWorkbook wb, XSSFRow row, short column, short halign, short valign)
  {
    XSSFCell cell = row.createCell(column);
    cell.setCellValue(new XSSFRichTextString("Align It"));
    CellStyle cellStyle = wb.createCellStyle();
    cellStyle.setAlignment(halign);
    cellStyle.setVerticalAlignment(valign);
    cell.setCellStyle(cellStyle);
  }

  private static void centerAcrossSelection(XSSFWorkbook wb, XSSFRow row, short start_column, short end_column, short valign)
  {
    XSSFCellStyle cellStyle = wb.createCellStyle();
    cellStyle.setAlignment((short)6);
    cellStyle.setVerticalAlignment(valign);

    for (int i = start_column; i <= end_column; i++) {
      XSSFCell cell = row.createCell(i);
      cell.setCellStyle(cellStyle);
    }

    XSSFCell cell = row.getCell(start_column);
    cell.setCellValue(new XSSFRichTextString("Align It"));

    CTRowImpl ctRow = (CTRowImpl)row.getCTRow();

    Object span = start_column + ":" + end_column;

    List spanList = new ArrayList();
    spanList.add(span);

    ctRow.setSpans(spanList);
  }
}