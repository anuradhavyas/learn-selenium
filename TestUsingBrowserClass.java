package test;

import java.io.FileOutputStream;
import java.util.List;

import org.openqa.selenium.WebElement;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import Core_Framework.Browser;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

//public class Write_Excel_Example_Anu{
public class TestUsingBrowserClass {

	WritableSheet ws;
	WritableSheet ws1;
	WritableWorkbook wb;

	@BeforeTest
	public void createExcel() throws Exception {
		wb = Workbook.createWorkbook(new FileOutputStream("/Users/Anu/Documents/Selenium/example.xls"));
		ws = wb.createSheet("WebPage_Links", 0);
		ws1 = wb.createSheet("TEstData", 1);
	}

	@Test
	public void chkExcelWriting() throws Exception {

		Browser.go("http://www.spicejet.com");

		// wd.get("http://www.spicejet.com");
		Thread.sleep(5000);

		// wd.findElement(By.xpath(".//*[@id='ctl00_mainContent_btn_FindFlights']")).click();
		Browser.click("id=ctl00_mainContent_btn_FindFlights");
		Browser.captureScreen("SpicejetHomepage");
		Thread.sleep(3000);
		// String alertMSG = wd.switchTo().alert().getText();
		// System.out.println(alertMSG);
		// wd.switchTo().alert().accept();//Accept to click on OK Button of
		// Alert
		// wd.switchTo().alert().dismiss();//Dismiss to click on Cancel Button
		// of ALert

		// wd.findElement(By.xpath(".//*[@id='ctl00_mainContent_ddl_originStation1_CTXTaction']")).click();
		Browser.click("id=ctl00_mainContent_ddl_originStation1_CTXTaction");
		List<WebElement> myvalues = Browser.findElements(".//*[@id='dropdownGroup1']/div/ul/li/a");
		// System.out.println(myvalues);
		// System.out.println(myvalues.size());
		int i = 0;
		for (WebElement elem : myvalues) {
			System.out.println(elem.getText());
			ws.addCell(new Label(0, i, elem.getText()));
			i++;
		}

		// for(String s:xyz)
		wb.write();
		wb.close();

		Browser.click("id=ctl00_mainContent_rbtnl_Trip_1");
		// clickin
		// on
		// one
		// way
		// wd.findElement(By.xpath(".//*[@id='ctl00_mainContent_ddl_originStation1_CTXTaction']")).click();
		Browser.click("id=ctl00_mainContent_ddl_originStation1_CTXTaction");
		Browser.click("xpath=.//*[@id='dropdownGroup1']/div/ul[1]/li[1]/a");
		// wd.findElement(By.xpath(".//*[@id='dropdownGroup1']/div/ul[1]/li[1]/a")).click();
		Thread.sleep(3000);
		Browser.click("id=ctl00_mainContent_ddl_destinationStation1_CTXTaction");
		// wd.findElement(By.xpath(".//*[@id='ctl00_mainContent_ddl_destinationStation1_CTXTaction']")).sendKeys("BLR");
		// wd.findElement(By.xpath(".//*[@id='dropdownGroup1']/div/ul[1]/li[1]/a")).click();
		// wd.findElement(By.xpath(".//*[@id='ctl00_mainContent_view_date1']")).click();
		Browser.click("id=ctl00_mainContent_view_date1");
		Browser.click("xpath=.//*[@id='ui-datepicker-div']/div[1]/table/tbody/tr[5]/td[1]/a");
		// wd.findElement(By.xpath(".//*[@id='ui-datepicker-div']/div[1]/table/tbody/tr[3]/td[4]/a")).click();

		Browser.select("id=ctl00_mainContent_DropDownListCurrency", "USD");
		// Select currency = new
		// Select(wd.findElement(By.xpath(".//*[@id='ctl00_mainContent_DropDownListCurrency']")));
		// currency.selectByVisibleText("USD");

		Browser.click("id=ctl00_mainContent_btn_FindFlights");
		// wd.findElement(By.xpath(".//*[@id='ctl00_mainContent_btn_FindFlights']")).click();

		Thread.sleep(3000);
		// select currency converter link
		// wd.findElement(By.xpath(".//*[@id='popUpConverter']")).click();

		/*
		 * Select from_currency = new Select(wd.findElement(By.xpath(
		 * ".//*[@id='CurrencyConverterCurrencyConverterView_DropDownListBaseCurrency']"
		 * ))); from_currency.selectByVisibleText("Indian Rupees(INR)");
		 * 
		 * Select To_Currency = new Select(wd.findElement(By.xpath(
		 * ".//*[@id='CurrencyConverterCurrencyConverterView_DropDownListConversionCurrency']"
		 * ))); To_Currency.selectByVisibleText("US Dollar(USD)");
		 * 
		 * wd.findElement(By.xpath(
		 * ".//*[@id='CurrencyConverterCurrencyConverterView_TextBoxAmount']")).
		 * sendKeys("100");//entering covnersion amount wd.findElement(By.xpath(
		 * ".//*[@id='CurrencyConverterCurrencyConverterView_ButtonConvert']")).
		 * click();//click on convert button
		 * wd.findElement(By.xpath(".//*[@id='ButtonCloseWindow']")).click();//
		 * click on Close Button
		 */
		Browser.captureScreen("Search Results");

	}

}
