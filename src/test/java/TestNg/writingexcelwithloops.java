package TestNg;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writingexcelwithloops {

	public static void main(String[] args) throws IOException {
		Scanner sc = new Scanner(System.in);
		
		 FileOutputStream File = new FileOutputStream("C:\\Users\\karna\\Downloads\\java project\\tcs\\TestData\\data05.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook();
	         XSSFSheet Sheet= workbook.createSheet();
	         for(int r=0; r<3;r++)
	         {
	        	 XSSFRow currentrow = Sheet.createRow(r);
	        	 
	        	 for (int c=0;c<3;c++)
	        	 {
	        		 System.out.println("enter the value:");
	        		 String values = sc.next();
	                 currentrow.createCell(c).setCellValue(values);
	                 
	           
	        		
	        	 }
	         }
	         
              workbook.write(File);
              workbook.close();
              File.close();
              
	}

}
