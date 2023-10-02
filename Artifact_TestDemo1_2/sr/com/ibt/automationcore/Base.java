package com.ibt.automationcore;

import java.util.NoSuchElementException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.WebDriverWait;

import freemarker.log.Logger;


public class Base {
	public static WebDriver driver;
	public static WebDriverWait wait; 

	protected static Logger log = Logger.getLogger(Base.class.getName()); 
//	@BeforeMethod
/*   public static void initiateDriverInstance(String browserName) throws Exception {
	   if(Utility.fetchinvokedetails("browserName").toString().equalsIgnoreCase("chrome"))
	   {
		System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir")+"/Drivers/chromedriver.exe");
//		System.setProperty("webdriver.chrome.driver", "D:\\maven automation programs\\AutomationTest_Artifact1_Instagram\\Drivers\\chromedriver.exe");   
	    driver=new ChromeDriver();
	    
	   }
	   else if(Utility.fetchinvokedetails("browserName").toString().equalsIgnoreCase("firefox"))
	   {
		   driver=new FirefoxDriver();
	   }
	   
	// driver.get(Utility.fetchinvokedetails("appURL").toString());  
	   Thread.sleep(4000); 
	   driver.quit(); 
	  
   } */
	public static void launchBrowser(String browser) {	 	
		 try{ 
			if (browser.equalsIgnoreCase("IE")) {  
				log.info("IE Browser is launching...");
				System.setProperty("webdriver.ie.driver",System.getProperty("user.dir")+"/sr/drivers/IEDriverServer.exe");
			    // create IE instance
				driver = new InternetExplorerDriver(); 
				log.info("IE Browser is launched.");
	//			implicitlywait();		                  
     			
		  } else if (browser.equalsIgnoreCase("Chrome")) { 
				log.info("Chrome Browser is launching...");
			 	System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir")+"/sr/drivers/chromedriver.exe");
//				System.setProperty("webdriver.chrome.driver", "D:\\maven automation programs\\AutomationTest_Artifact1_Instagram\\Drivers\\chromedriver.exe");   
			 	
               driver = new ChromeDriver();          
			}   
		 }
	  catch(Exception e){
		 System.out.println("Failed to launch Browser.");
	 }
		 }               

   /*	@AfterMethod
   public void closeDriverInstance() {
	   driver.quit();
   } */
	public static void sendURL(String url) {
		log.info("application url is being opened...");
		driver.navigate().to(url);
		log.info("application url is opened.");
	}
	public static void sendValue(String locator, String testdata) {

		try {	
			if(!isElementVisible(locator)) {
			    return;
			}

			driver.findElement(By.xpath(locator)).clear();
			log.info("Cleared textbox data");
			driver.findElement(By.xpath(locator)).sendKeys(testdata);
			log.info("Data sent to textbox");
		} catch (NoSuchElementException e) {
			e.printStackTrace();
			log.info("Failed to send the data to textbox");
		}
	}
public static void rightClick(String source) {
		
		try {
			WebElement elementToRightClick = driver.findElement(By.xpath(source));
			Actions clicker = new Actions(driver);
			clicker.contextClick(elementToRightClick).perform();
			log.info("Right click operation performed");
		} catch (Exception e) {
			log.info("Unable to perform right click operation");
		} 
	}

public static void click(String locator){
	try {
		if(!isElementVisible(locator)) {
		    return;
		}
		driver.findElement(By.xpath(locator)).click();
		log.info("Click operation done.");
	} catch (Exception e) {
		System.out.println(e.getMessage());
		log.info("Click operation failed.");
	}
}


public static  boolean  isElementVisible(String locator) {
    try {
        if (driver.findElement(By.xpath(locator)).isDisplayed()) {
            log.info("Element is Displayed");
            return true;
        }
    }
    catch(Exception e) {       
        log.info("Element is Not Displayed");
        return false;
    }       
    return false;
}

	
}
