package com.seetech.automation.base;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.ITestContext;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Parameters;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.seetech.automation.utility.ConfigManager;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseTest {

	public static WebDriver driver;
	public static ExtentReports extentReports;
	public static ExtentTest extentTest;
	public static String startDate;
	public static String reportsDestination;
	public String reportFilePath;
	
	@BeforeSuite
	public void beforeAll(ITestContext itc) {
		
		SimpleDateFormat sdf = new SimpleDateFormat("MMM_dd_yyyy_z_HH_mm_ss");
		startDate = sdf.format(new Date());
		reportFilePath = System.getProperty("user.dir")+File.separator+"Reports"+File.separator+"reports_"+startDate+".html";
		extentReports = new ExtentReports(reportFilePath);
	}

	@Parameters({ "browser" })
	@BeforeMethod
	public void setup(ITestContext itc, String browser) {

		if ("chrome".equals(browser)) {
			WebDriverManager.chromedriver().setup();
			driver = new ChromeDriver();
		} else if("ie".equals(browser)) {
			WebDriverManager.iedriver().setup();
			driver = new InternetExplorerDriver();
		} else if("firefox".equals(browser)) {
			WebDriverManager.firefoxdriver().setup();
			driver = new FirefoxDriver();
		}
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		driver.get(ConfigManager.getProperty("url"));

	}

	@AfterMethod
	public void rearDown() {
		driver.close();
		extentReports.flush();
	}
	
	/**
	 * This method open the Extent report generated in a default browser after suite
	 * execution completed.
	 * 
	 */
	@AfterSuite
	public void SuiteClose() {
		File file = new File(reportFilePath);
		Desktop desktop = Desktop.getDesktop();
		try {
			desktop.browse(file.toURI());
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
