import java.io.File;
import java.io.IOException;

import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
public class Excel1 {
	 @DataProvider(name = "DP1")
	public Object[][] cool()
{
		 String[][] tabArray=new String[2][3];
	File file = new File("C:\\Users\\ankitmalik\\new1.ods");
	Sheet sheet;
	try {
		// Getting the 0th sheet for manipulation| pass sheet name as string
		sheet = SpreadSheet.createFromFile(file).getSheet(0);

		// Get row count and column count
		int nRowCount = sheet.getRowCount();
		int nColCount = sheet.getColumnCount();
		
		// Iterating through each row of the selected sheet
		MutableCell cell = null;
		for (int nRowIndex = 0; nRowIndex < nRowCount-1; nRowIndex++) {
			for (int nColIndex = 0; nColIndex < nColCount; nColIndex++){
			cell = sheet.getCellAt(nColIndex,nRowIndex);

			// System.out.println(cell.getValue()+" "+key);

			//if (key.equals(cell.getValue().toString())) 
			{
				cell = sheet.getCellAt(nColIndex,nRowIndex+1);
				 tabArray[nRowIndex][nColIndex] = cell.getValue().toString();
				 //System.out.println(cell.getValue().toString());
				 
			}
			}
		}
	} catch (IOException e) {
		e.printStackTrace();
	}
	return tabArray;	
}
	 @Test (dataProvider = "DP1")
	    public void testDataProviderExample(String ex1, 
	            String ex2, String ex3) throws Exception {    
	            System.setProperty("webdriver.chrome.driver", "D:\\chromedriver.exe");
	   	        WebDriver driver = new ChromeDriver();
	   	        driver.get("https://webmail.qainfotech.com:8443");
	   	        Thread.sleep(1000);
	   	        driver.findElement(By.id("username")).sendKeys("ankitmalik@qainfotech.com");
	   	        driver.findElement(By.id("password")).sendKeys("Malikankit@12");
	   	        driver.findElement(By.cssSelector("input.ZLoginButton.DwtButton")).click();
	   	        Thread.sleep(1000);
	   	        driver.findElement(By.id("zb__NEW_MENU_title")).click();
	   	        Thread.sleep(2000);
	   	        driver.findElement(By.cssSelector("input.addrInputField.user_font_system")).sendKeys(ex1);
	   	        driver.findElement(By.cssSelector("input.subjectField")).sendKeys(ex2);
	   	        driver.switchTo().frame("ZmHtmlEditor1_body_ifr");
	   	        driver.findElement(By.id("tinymce")).sendKeys(ex3);
	   	        driver.switchTo().defaultContent();
	   	        driver.findElement(By.id("zb__COMPOSE-1__SEND_title")).click();
	    }
}
