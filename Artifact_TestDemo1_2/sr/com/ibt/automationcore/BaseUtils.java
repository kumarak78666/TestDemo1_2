package com.ibt.automationcore;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.ibt.config.ApplicationConfig;
import com.ibt.utilities.ExcelReader;

import org.testng.annotations.BeforeMethod;
import org.testng.annotations.AfterMethod;
import com.ibt.utilities.Property;
import com.ibt.constants.Constants;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.*;
public class BaseUtils extends Base{
	
	static ExcelReader read;
	static Property prop = null;
	static Property adminProps = null;
	static Property finProps = null;
	static Property userProps = null;
	public static ExtentReports extent;
	public static ExtentTest expense;
	
	public static void  testSetup(){
		try {
			read=new ExcelReader(Constants.testDataPath);
			prop = new Property(Constants.objectRepoPath);
	/*		adminProps = new Property(Constants.adminProperties, Property.ADMIN_MODULE);
			finProps = new Property(Constants.financeProperties, Property.FINANCE_MODULE);
			userProps = new Property(Constants.userProperties, Property.USER_MODULE); */
	//		extent=new ExtentReports(Constants.reportPath, ApplicationConfig.appendReports); //set the report path and whether to append the report to previous report or not
	//		extent.loadConfig(new File(Constants.reportConfigPath)); //Configuration of extent reports			
		} catch (InvalidFormatException | IOException e) {
			e.printStackTrace();
		}
	}
	
	@BeforeSuite
	public void beforeSuite(){
		try {
			testSetup();
	//		login();
		} catch (Exception e ) {
			e.printStackTrace();
		}
	}
	
	
	@AfterSuite
	public void afterSuite(){
		try {
			logout();
			quitBrowser();
		} catch (Exception e ) {
			e.printStackTrace();
		}
	}

	
	
	private void quitBrowser() {
		// TODO Auto-generated method stub
		
	}

	private void logout() {
		// TODO Auto-generated method stub
		
	}

	@Test
	public static void login() throws InterruptedException{ 
		
		
		try{
		/*	if(!Thread.currentThread().getStackTrace()[2].getMethodName().equalsIgnoreCase("afterSuite")) {
				return;   
			}*/
			launchBrowser(ApplicationConfig.browserName);
	 		log.info("chrome Browser is launched.");                         
		//	windowMaximize();
			sendURL(ApplicationConfig.appURL);      
			log.info("link opened");   
			System.out.println("link opened");
			ExcelReader.setTestCaseNameRowNo(1); 
			System.out.println("click on excel");
			Thread.sleep(2000);
		//	expense=extent.startTest(testName + " with "+ExcelReader.getCellData("Login", "UserName"));

			sendValue(Property.getProperty("userName"),ExcelReader.getCellData("Login", "UserName"));
		//	click(Property.getProperty("userName"));
		//	 sendValue(Property.getProperty("userName"), "test@ibt.com");
		        log.info("clicked on propertyelement");  
			System.out.println("User Name entered");
		//	expense.log(LogStatus.PASS, "User Name entered");
			sendValue(Property.getProperty("password"), ExcelReader.getCellData("Login", "Password"));
		//	expense.log(LogStatus.PASS, "Password entered");
			click(Property.getProperty("signIn"));
		//	expense.log(LogStatus.PASS, "Clicked on Sign In");
			Thread.sleep(1500);

		} catch (Exception e) { 
			e.printStackTrace();
		}
	}


}