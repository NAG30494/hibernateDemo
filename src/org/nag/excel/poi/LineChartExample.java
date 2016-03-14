package org.nag.excel.poi;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.ChartLegend;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LineChartExample {

	public static void drawChat(Sheet sheet, int startColumn, int startRow,	int endColumn, int endRow, int[][] cellRanges) {
		int[] values;
		ChartDataSource chartDataSource = null;
		Drawing drawing = sheet.createDrawingPatriarch();
		ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, startColumn,startRow, endColumn, endRow);

		Chart chart = drawing.createChart(anchor);
		ChartLegend legend = chart.getOrCreateLegend();
		legend.setPosition(LegendPosition.BOTTOM);

		LineChartData data = chart.getChartDataFactory().createLineChartData();

		ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
		ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
		leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

		ChartDataSource source[] = new ChartDataSource[cellRanges.length];

		for (int i = 0; i < cellRanges.length; i++) {
			values = cellRanges[i];
			source[i] = DataSources.fromNumericCellRange(sheet,
					new CellRangeAddress(values[0], values[1], values[2],
							values[3]));
		}
		chartDataSource = source[0];
		for (int i = 1; i < source.length; i++) {
			data.addSeries(chartDataSource, source[i]);
		}
		chart.plot(data, new ChartAxis[] { bottomAxis, leftAxis });
	}

	public static void main(String[] args) throws Exception {
		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("linechart");
		int NUM_OF_ROWS = 10;
		int NUM_OF_COLUMNS = 4;
		Row row = null;
		int[] value = null;
		String names[] = { "Seller ", "Peter Jackson", "Phil Smith",
				"Joseph Cole" };
		int values[][] = { { 2003, 80000, 56000, 99000 },
				{ 2004, 75000, 69000, 58000 }, { 2005, 68000, 91000, 78000 },
				{ 2006, 91000, 86000, 81000 }, { 2007, 78000, 88000, 71000 },
				{ 2008, 83000, 90000, 78000 }, { 2009, 85000, 68000, 57000 },
				{ 2010, 90000, 83000, 78000 }, { 2011, 79000, 85000, 90000 } };
		row = sheet.createRow(0);
		for (int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++) {
			Cell cell = row.createCell((short) colIndex);
			cell.setCellValue(names[colIndex]);
		}

		for (int rowIndex = 1; rowIndex < NUM_OF_ROWS; rowIndex++) {
			row = sheet.createRow((short) rowIndex);
			value = values[rowIndex - 1];
			for (int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++) {
				Cell cell = row.createCell((short) colIndex);
				cell.setCellValue(value[colIndex]);
			}
		}
		// firstRow, lastRow, firstCol, lastCol
		int[][] cellRanges = { { 0, 9, 0, 0 }, { 0, 9, 1, 1 }, { 0, 9, 2, 2 },
				{ 0, 9, 3, 3 } };
		drawChat(sheet, 0, 15, 10, 29, cellRanges);
		drawChat(sheet, 11, 5, 20, 15, cellRanges);

		FileOutputStream fileOut = new FileOutputStream("ooxml-line-chart.xlsx");
		wb.write(fileOut);
		fileOut.close();
	}
}
