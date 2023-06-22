package TestNg;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reading {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		FileInputStream file= new FileInputStream("C:\\Users\\kanna\\eclipse-workspace\\TataMotors\\TestData\\Data.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet Sheet=workbook.getSheet("Sheet1");
		
		int totalrows= Sheet.getLastRowNum();
		int totalcells=Sheet.getRow(1).getLastCellNum();
		System.out.println(totalrows);
		System.out.println(totalcells);
	   
		
		for(int r=0;r<totalrows;r++)
		{
			XSSFRow currentRow=Sheet.getRow(r);
			for(int c=0;c<totalcells;c++)
			{
				String values=currentRow.getCell(c).toString();
				System.out.print(values+" ");
			}
			System.out.println();


			}


	}

}
