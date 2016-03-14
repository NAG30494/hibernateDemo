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

public class LineChart {
	
	
	 public static void drawChat( Sheet sheet,int startColumn, int startRow,int  endColumn, int endRow,int [][] cellRanges ){
		 int chartStartColumn;
		 int chartEndColumn;
		 int chartStartRow;
		 int chartEndRow;
		 int[] values;
		 ChartDataSource c= null;
		 Drawing drawing = sheet.createDrawingPatriarch();
		 //ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 10, 15);
		    ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, startColumn, startRow, endColumn, endRow);

		    Chart chart = drawing.createChart(anchor);
		    ChartLegend legend = chart.getOrCreateLegend();
		    //legend.setPosition(LegendPosition.TOP_RIGHT);
		    legend.setPosition(LegendPosition.BOTTOM);

		    LineChartData data = chart.getChartDataFactory().createLineChartData();

		    ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
		    ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
		    leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

		    ChartDataSource source[] =new ChartDataSource[cellRanges.length];
		    
		    for(int i=0;i<cellRanges.length;i++ ){
		    	values= cellRanges[i];
		    	chartStartColumn=values[0];
		    	chartStartRow=values[1];
		    	chartEndColumn=values[2];
		    	chartEndRow=values[3];
		    	
		    	//c = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, 9));
		    	source[i]=DataSources.fromNumericCellRange(sheet, new CellRangeAddress(chartStartColumn, chartStartRow, chartEndColumn, chartEndRow));
		    	
		    }

		    c=source[0];
		    for(int i=1;i<source.length;i++){
		    	data.addSeries(c,source[i]);
		    }
		    
		    

		    chart.plot(data, new ChartAxis[] { bottomAxis, leftAxis });
	 }
	public static void main(String[] args)
		    throws Exception
		  {
		    Workbook wb = new XSSFWorkbook();
		    Sheet sheet = wb.createSheet("linechart");
		    int NUM_OF_ROWS = 3;
		    int NUM_OF_COLUMNS = 10;

		    for (int rowIndex = 0; rowIndex < 3; rowIndex++) {
		      Row row = sheet.createRow((short)rowIndex);
		      for (int colIndex = 0; colIndex < 10; colIndex++) {
		        Cell cell = row.createCell((short)colIndex);
		        cell.setCellValue(colIndex * (rowIndex + 1));
		      }
		    }
		    
		    int[][] cellRanges ={{0, 0, 0, 9},{1, 1, 0, 9},{2, 2, 0, 9}};
		    
		    drawChat(sheet, 0, 5, 10, 15, cellRanges);
		    drawChat(sheet, 11, 5, 20, 15, cellRanges);
		   
		    
		    

		    FileOutputStream fileOut = new FileOutputStream("ooxml-line-chart.xlsx");
		    wb.write(fileOut);
		    fileOut.close();
		  }
}
