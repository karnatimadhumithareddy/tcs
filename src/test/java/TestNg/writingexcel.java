package TestNg;



import java.io.FileOutputStream;
import java.io.IOException;


import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writingexcel {

	public static void main(String[] args) throws IOException {
		
		 FileOutputStream File = new FileOutputStream("C:\\Users\\karna\\Downloads\\java project\\tcs\\TestData\\data06.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook();
         XSSFSheet Sheet= workbook.createSheet();
         
         XSSFRow row1 = Sheet.createRow(0);
         row1.createCell(0).setCellValue("Name");
         row1.createCell(1).setCellValue("city");
         row1.createCell(2).setCellValue("job");
         
         XSSFRow row2 = Sheet.createRow(1);
         row2.createCell(0).setCellValue("madhu");
         row2.createCell(1).setCellValue("delhi");
         row2.createCell(2).setCellValue("teacher");
         
         workbook.write(File);
         workbook.close();
         File.close();
         
         System.out.println("Done");
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
	}
	

}
