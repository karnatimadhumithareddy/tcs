   package TestNg;

import java.io.IOException;
import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class datadriventesting {

	@Test
	void All() throws IOException {
		WebDriver driver = new ChromeDriver();
		driver.get("https://www.cit.com/cit-bank/resources/calculators/certificate-of-deposit-calculator");
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		
		WebElement InitialDeposit = driver.findElement(By.id("mat-input-0"));
	    WebElement Length = driver.findElement(By.id("mat-input-1"));
		WebElement InterestRate = driver.findElement(By.id("mat-input-2"));
		WebElement Calbutton = driver.findElement(By.xpath("//*[@id=\"CIT-chart-submit\"]/div"));
		
		
		System.out.println("Identified elements to calculate CD");
		
		String path = "C:\\Users\\karna\\Downloads\\java project\\tcs\\TestData\\datadrivendata.xlsx";
		int rows= ExcelUtils.getRowCount(path, "Sheet1");
		

   System.out.println("Row counts:"+rows);
   for(int i=1; i<rows;i++)
   {
   String inidep=ExcelUtils.getcelldata(path, "Sheet1", i, 0);
   String interestrate=ExcelUtils.getcelldata(path, "Sheet1", i, 1);
   String monthlength=ExcelUtils.getcelldata(path, "Sheet1", i, 2);
   //String Compoundingmonths=ExcelUtils.getcelldata(path, "Sheet1", i, 3);
   String exptotal=ExcelUtils.getcelldata(path, "Sheet1", i, 4);
   
   InitialDeposit.clear();
   Length.clear();
   InterestRate.clear();
   
   InitialDeposit.sendKeys(inidep);
   Length.sendKeys(monthlength);
   InterestRate.sendKeys(interestrate);
   Calbutton.click();


   String acttotal=driver.findElement(By.xpath("//*[@id=\"displayTotalValue\"]")).getText();
   System.out.println(acttotal);
   
   if(exptotal.equals(acttotal))
   {
   ExcelUtils.Setcelldata(path, "Sheet1", i, 6, "passed");
   ExcelUtils.fillGreenColor(path, "Sheet1", i, 6);
   }
   else
   {
   ExcelUtils.Setcelldata(path,"Sheet1",i, 6,"Failed");
   ExcelUtils.fillRedColour(path, "Sheet1", i, 6);
   }
   }
   System.out.println("Done");
   }

		
		

	}


