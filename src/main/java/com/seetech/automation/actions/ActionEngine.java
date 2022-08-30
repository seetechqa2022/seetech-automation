package com.seetech.automation.actions;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import com.relevantcodes.extentreports.LogStatus;
import com.seetech.automation.base.BaseTest;

public class ActionEngine extends BaseTest {

	
	public void click(By locator, String locatorName) throws Throwable{
		boolean flag =false;
		try {
		driver.findElement(locator).click();
		flag = true;
		}catch(Exception e) {
			e.printStackTrace();
		}finally{
			if(flag) {
				extentTest.log(LogStatus.PASS, "Succfully clicked on locator "+locatorName);
			}else {
				extentTest.log(LogStatus.FAIL, "Failed to click on locator "+locatorName + extentTest.addScreenCapture(getScreenshot(locatorName)));
			}
		}
	
	}
	
	public String getText(By locator, String locatorName) throws Throwable{
		boolean flag = false;
		String text="";
		try {
		 text = driver.findElement(locator).getText();
		flag = true;
	}catch(Exception e) {
		e.printStackTrace();
	}finally{
		if(flag) {
			extentTest.log(LogStatus.PASS, "Extracted text from the locator "+locatorName);
		}else {
			extentTest.log(LogStatus.FAIL, "Failed to extract text from the locator "+locatorName+ extentTest.addScreenCapture(getScreenshot(locatorName)));
		}
	}
		return text;
	}
	
	public void type(By locator, String data, String locatorName) throws Throwable{
		boolean flag =false;
		try {
		driver.findElement(locator).sendKeys(data);
		flag = true;
		}catch(Exception e) {
			e.printStackTrace();
		}finally{
			if(flag) {
				extentTest.log(LogStatus.PASS, "Succfully entered given text into "+locatorName);
			}else {
				extentTest.log(LogStatus.FAIL, "Failed to entered given text into "+locatorName+ extentTest.addScreenCapture(getScreenshot(locatorName)));
			}
		}
	}
	
	public static String getScreenshot(String screenshotName) throws Throwable {
		String screenshotLocation = System.getProperty("user.dir");
		try {
			String dateName = new SimpleDateFormat("yyyyMMddhhmmss").format(new Date());
			TakesScreenshot ts =(TakesScreenshot)driver;
			File source = ts.getScreenshotAs(OutputType.FILE);
			screenshotLocation= screenshotLocation+File.separator+ "FailedScreenShots"+File.separator + screenshotName+ dateName + ".png";
			File finalDestination = new File(screenshotLocation);
			FileUtils.copyFile(source, finalDestination);
			
		}catch(Exception e) {
			e.printStackTrace();
		}
		
		return screenshotLocation;
	}
	
}
