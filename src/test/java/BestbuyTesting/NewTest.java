package BestbuyTesting;

import java.util.regex.Pattern;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import java.util.Set;
import java.util.concurrent.TimeUnit;


import java.io.File;


import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.Workbook;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.chrome.ChromeDriver;

import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Listeners;

import atu.testng.reports.ATUReports;
import atu.testng.reports.listeners.ATUReportsListener;
import atu.testng.reports.listeners.ConfigurationListener;
import atu.testng.reports.listeners.MethodListener;

import atu.testng.reports.logging.LogAs;
import atu.testng.selenium.reports.CaptureScreen;
import atu.testng.selenium.reports.CaptureScreen.ScreenshotOf;

//import jxl.Sheet;
//import jxl.Workbook;

@Listeners({ ATUReportsListener.class, ConfigurationListener.class, MethodListener.class})
public class NewTest {

	private WebDriver driver;
	private WebDriverWait myWaitVar;
	private String baseUrl="https://www.bestbuy.com/";
	private Sheet  addressSheet;
	private Sheet  creditCard;
	
	@BeforeTest
	public void setUp() throws Exception {
		//Reading Excel file
		//reference http://viralpatel.net/blogs/java-read-write-excel-file-apache-poi/
		FileInputStream fs = new FileInputStream("C:\\Users\\1323928\\workplace\\WebdriverTest\\shipping_address.xls");
	    Workbook wb = new HSSFWorkbook(fs);
	    addressSheet = wb.getSheet("Sheet1");
	    
	    for(int i=0;i<=addressSheet.getLastRowNum();i++){
	    	System.out.println("row Number: "+i);
	    	for(int j=0;j<addressSheet.getRow(i).getLastCellNum();j++){
	    		System.out.println(addressSheet.getRow(i).getCell(j));
	    	}
	    }
	    
	    HSSFWorkbook NewWorkbook = new HSSFWorkbook();
		HSSFSheet sheet = NewWorkbook.createSheet("Sample sheet");
		for(int i=0;i<=addressSheet.getLastRowNum();i++){
	    	System.out.println("row Number: "+i);
	    	sheet.createRow(i);
	    	for(int j=0;j<addressSheet.getRow(i).getLastCellNum();j++){
	    		System.out.println(addressSheet.getRow(i).getCell(j));
	    		Cell cell=addressSheet.getRow(i).getCell(j);
	    		//sheet.getRow(i).
	    	}
	    }
		
		FileOutputStream out = new FileOutputStream(new File("C:\\Users\\1323928\\workplace\\WebdriverTest\\NewWorkbook.xls"));
		NewWorkbook.write(out);
		out.close();
	    //Workbook wb2 = Workbook.getWorkbook(fs2);
	    //creditCard=wb2.getSheet(0);
	    
			System.setProperty("atu.reporter.config", "C:\\Users\\1323928\\Desktop\\selenium\\atu.properties");
			System.setProperty("atu.reports.takescreenshot", "true");
			System.setProperty( "webdriver.chrome.driver", "C:\\Users\\1323928\\Desktop\\selenium\\chromedriver.exe" );
			driver = new ChromeDriver();
			driver.get(baseUrl);
			driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);

			myWaitVar=new WebDriverWait(driver,10);

			ATUReports.setWebDriver(driver);
			ATUReports.indexPageDescription = "Test Project";
		    FileInputStream fs2 = new FileInputStream("C:\\Users\\1323928\\workplace\\WebdriverTest\\creditcardinfo.xls");

			
		    //System.setProperty( "webdriver.gecko.driver", "C:\\Users\\1323928\\Downloads\\geckodriver.exe");		  
		    //WebDriver driver = new FirefoxDriver();
	}
	
	@AfterTest
	public void testWebdriver() throws Exception {

		driver.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);
		//get rid of the pop up web elements
		try{

			if(driver.findElement(By.id("emailModalLabel")).isDisplayed()){			
				driver.findElement(By.cssSelector("button.close")).click();
			};
		}catch(org.openqa.selenium.NoSuchElementException e){
			   System.out.println(e);
		}	
		driver.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);
		try{		
			
			if(driver.findElement(By.id("survey_invite_no")).isDisplayed()){
				  driver.findElement(By.id("survey_invite_no")).click();
			}
		}catch(org.openqa.selenium.NoSuchElementException e){
		   System.out.println(e);
		}
        
		//search an item and change its quantity in cart
		changequantity();
		
		/*
		//search and add 2 items to cart
		searchandadd2itemstocart();

		//go to checkout page
		checkoutPage();
		
		//filling out the shipping address
		shippingPage();	
		
		//filling out credit card info
		validateCreditcard();
		//ATUReports.add("INfo Step", LogAs.INFO, new CaptureScreen(ScreenshotOf.BROWSER_PAGE));
		 * 
		 */
	}

	@AfterTest
	public void tearDown() throws Exception {
		//driver.quit();
	}
	
	public void validateCreditcard() throws InterruptedException{
		Thread.sleep(2000);
		driver.findElement(By.id("creditCard")).sendKeys("2338472818237777");
		Thread.sleep(2000);
		driver.findElement(By.id("securityCode")).sendKeys("117");
		new Select(driver.findElement(By.id("expirationMonth"))).selectByVisibleText("06 - June");
		new Select(driver.findElement(By.id("expirationYear"))).selectByVisibleText("2019");
		driver.findElement(By.id("emailAddress")).sendKeys("jinyan.li@tcs.com");
		driver.findElement(By.cssSelector("button.place-order")).click();
		Thread.sleep(2000);
		driver.findElement(By.cssSelector("span.help-block__text")).isDisplayed();
		//ATUReports.add("Warning Step", LogAs.WARNING,new CaptureScreen(driver.findElement(By.cssSelector("span.help-block__text"))));

	}
	
	public void shippingPage() throws InterruptedException{
		myWaitVar.until(ExpectedConditions.visibilityOf(driver.findElement(By.className("shipping-address-btn")))).click();
		myWaitVar.until(ExpectedConditions.visibilityOf(driver.findElement(By.name("firstName")))).clear();
		/*
        driver.findElement(By.name("firstName")).sendKeys(addressSheet.getCell(0, 1).getContents());
        driver.findElement(By.name("lastName")).sendKeys(addressSheet.getCell(1, 1).getContents());
        driver.findElement(By.name("street")).sendKeys(addressSheet.getCell(2, 1).getContents());
        driver.findElement(By.name("city")).sendKeys(addressSheet.getCell(3, 1).getContents());
        */
        new Select(driver.findElement(By.name("state"))).selectByVisibleText("OH - Ohio");
        /*
        driver.findElement(By.name("zipcode")).sendKeys(addressSheet.getCell(5, 1).getContents());
        driver.findElement(By.name("dayPhoneNumber")).sendKeys(addressSheet.getCell(6, 1).getContents());
        */
        driver.findElement(By.cssSelector("input.btn.btn-secondary")).click();
        Thread.sleep(2000);
        myWaitVar.until(ExpectedConditions.visibilityOf(driver.findElement(By.className("continue")))).click();
	}
	
	public void checkoutPage() throws InterruptedException{
		driver.findElement(By.cssSelector("span.header-icon-cart")).click();
		Thread.sleep(5000);
		driver.findElement(By.xpath("(//input[@name='fulfillmentType'])[2]")).click();
		Thread.sleep(2000);
		driver.findElement(By.xpath("(//input[@name='fulfillmentType'])[4]")).click();
		Thread.sleep(2000);
		driver.findElement(By.linkText("Checkout")).click();
		Thread.sleep(2000);
		myWaitVar.until(ExpectedConditions.visibilityOf(driver.findElement(By.linkText("Continue as Guest")))).click();
		
	}
	
	public void changequantity() throws InterruptedException{
		driver.findElement(By.id("gh-search-input")).clear();
		driver.findElement(By.id("gh-search-input")).sendKeys("mad max fury road");
		driver.findElement(By.cssSelector("button.header-search-button")).click();
		Thread.sleep(2000);
		myWaitVar.until(ExpectedConditions.elementToBeClickable(driver.findElement(By.cssSelector("span.label-add-to-cart")))).click();
		driver.manage().timeouts().implicitlyWait(3, TimeUnit.SECONDS);

		//get rid of the product suggestion pop up 
		closePopUp();
		
		//go to cart page
		//myWaitVar.until(ExpectedConditions.elementToBeClickable(driver.findElement(By.cssSelector("span.header-icon-cart")))).click();
		//driver.findElement(By.linkText("Cart"));
		
		driver.get("http://www.bestbuy.com/cart");
		//update quantity
		driver.findElement(By.className("qty-input")).click();
		driver.findElement(By.className("qty-input")).clear();
		Thread.sleep(2000);
		driver.findElement(By.className("qty-input")).sendKeys("2");
		Thread.sleep(2000);
		driver.findElement(By.className("update-link")).click();
		driver.findElement(By.className("qty-input")).clear();
		driver.findElement(By.className("qty-input")).sendKeys("1");
		Thread.sleep(2000);
		driver.findElement(By.linkText("Remove")).click();
		
	}
	
	public void searchandadd2itemstocart() throws InterruptedException{
		driver.findElement(By.id("gh-search-input")).clear();
		driver.findElement(By.id("gh-search-input")).sendKeys("captain america civil war");
		driver.findElement(By.cssSelector("button.search-button")).click();
		driver.manage().timeouts().implicitlyWait(3, TimeUnit.SECONDS);
		driver.findElement(By.cssSelector("span.label-add-to-cart")).click();
		Thread.sleep(2000);
		closePopUp();
		driver.findElement(By.id("gh-search-input")).clear();
		driver.findElement(By.id("gh-search-input")).sendKeys("Batman v Superman");
		driver.findElement(By.cssSelector("button.search-button")).click();
		driver.manage().timeouts().implicitlyWait(3, TimeUnit.SECONDS);
		driver.findElement(By.cssSelector("span.label-add-to-cart")).click();
		Thread.sleep(2000);
		closePopUp();

	}
	
	public void closePopUp() throws InterruptedException{	
		Thread.sleep(2000);
		try{
			myWaitVar.until(ExpectedConditions.visibilityOf(driver.findElement(By.cssSelector("div.close-icon")))).click();
		}catch(org.openqa.selenium.NoSuchElementException e){
			   System.out.println(e);
		}
		
	}

}
