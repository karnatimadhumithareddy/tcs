package TestNg;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.opentelemetry.sdk.metrics.data.Data;

public class ExcelUtils {
	public static FileInputStream fi;
	public static FileOutputStream fo;
	public static XSSFWorkbook wb;
	public static XSSFSheet ws;
	public static XSSFRow row;
	public static XSSFCell cell;
	public static CellStyle Style;
	
	public static int getRowCount(String Xlfile, String XlSheet) throws IOException
	{
	fi=new FileInputStream(Xlfile);
	wb=new XSSFWorkbook(fi);
	ws= wb.getSheet(XlSheet);
	int rowcount= ws.getLastRowNum();
	wb.close();
	fi.close();
	return rowcount;
	}
	public static int getcellcount(String Xlfile, String XlSheet,int rownum) throws IOException
	{
	fi=new FileInputStream(Xlfile);
	wb=new XSSFWorkbook(fi);
	ws= wb.getSheet(XlSheet);
	row=ws.getRow(rownum);
	int cellcount= row.getLastCellNum();
	wb.close();
	return cellcount;
	}
	public static String getcelldata(String Xlfile, String XlSheet,int rownum,int column) throws IOException
	{
	fi=new FileInputStream(Xlfile);
	wb=new XSSFWorkbook(fi);
	ws= wb.getSheet(XlSheet);
	row=ws.getRow(rownum);
	cell=row.getCell(column);
	String data;
	try
	{
	//data=cell.toString();
	DataFormatter formatter= new DataFormatter();
	data=formatter.formatCellValue(cell);
	return data;
	}
	catch(Exception e)
	{
	data=" ";
	}
	wb.close();
	fi.close();
	return data;
	}




	public static void Setcelldata(String Xlfile, String XlSheet,int rownum,int column, String data) throws IOException
	{
	fi=new FileInputStream(Xlfile);
	wb=new XSSFWorkbook(fi);
	ws= wb.getSheet(XlSheet);
	row=ws.getRow(rownum);
	cell=row.getCell(column);
	cell.setCellValue(data);
	fo=new FileOutputStream(Xlfile);
	wb.write(fo);
	wb.close();
	fi.close();
	fo.close();
	}
	
	public static void fillGreenColor(String Xlfile, String XlSheet,int rownum,int column) throws IOException
	{
	fi=new FileInputStream(Xlfile);
	wb=new XSSFWorkbook(fi);
	ws= wb.getSheet(XlSheet);
	row=ws.getRow(rownum);
	cell=row.getCell(column);
	Style=wb.createCellStyle();
	Style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
	Style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	cell.setCellStyle(Style);
	fo=new FileOutputStream(Xlfile);
	wb.write(fo);
	wb.close();
	fi.close();
	fo.close();
	}

	public static void fillRedColour(String Xlfile, String XlSheet,int rownum,int column) throws IOException
	{
	fi=new FileInputStream(Xlfile);
	wb=new XSSFWorkbook(fi);
	ws= wb.getSheet(XlSheet);
	row=ws.getRow(rownum);
	cell=row.getCell(column);
	Style=wb.createCellStyle();
	Style.setFillForegroundColor(IndexedColors.RED.getIndex());
	Style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	cell.setCellStyle(Style);
	fo=new FileOutputStream(Xlfile);
	wb.write(fo);
	wb.close();
	fi.close();
	fo.close();
	}






	
}
