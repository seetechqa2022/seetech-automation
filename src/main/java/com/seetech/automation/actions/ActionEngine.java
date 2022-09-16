package com.seetech.automation.actions;

import java.io.File;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.FluentWait;

import java.awt.Robot;
import java.awt.datatransfer.StringSelection;
import java.awt.Toolkit;
import java.awt.AWTException;
import java.awt.event.KeyEvent;

import com.google.common.base.Function;
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
			driver.findElement(locator).sendKeys(Keys.CLEAR);
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
	
	  public WebElement getWebElement(By locator) {
		    return driver.findElement(locator);
		  }
	  
	  public void jsClick(By locator, String locatorName) throws Throwable {
		    boolean flag = false;
		    try {
		      WebElement webElement = getWebElement(locator);
		      JavascriptExecutor executor = (JavascriptExecutor) driver;
		      executor.executeScript("arguments[0].click();", webElement);
		      flag = true;
		    } catch (Exception e) {
		      e.printStackTrace();
		    } finally {
		      if (flag) {
		    	  extentTest.log(LogStatus.PASS, "JavaScript click is sucessful on element " + locatorName);
		      } else
		    	  extentTest.log(LogStatus.FAIL, "JavaScript click is unsucessful on element " + locatorName
		            + extentTest.addScreenCapture(getScreenshot(locatorName)));
		    }
		  }
	
	  public boolean selectByIndex(By locator, int index, String locatorName) throws Throwable {
		    boolean isByIndexSelected = false;
		    try {
		      WebElement webElement = getWebElement(locator);
		      Select dropDown = new Select(webElement);
		      dropDown.selectByIndex(index);
		      Thread.sleep(5000);
		      isByIndexSelected = true;
		    } catch (Exception e) {
		      e.printStackTrace();
		    } finally {
		      if (isByIndexSelected) {
		    	  extentTest.log(LogStatus.PASS, " Option at index " + index
		            + " is selected from the drop down " + locatorName);
		      } else {
		    	  extentTest.log(LogStatus.FAIL,
		            "Option at index " + index + " is Not selected from the drop dpwn " + locatorName
		                + extentTest.addScreenCapture(getScreenshot(locatorName)));
		      }
		    }
		    return isByIndexSelected;
		  }
	  
	  public boolean selectByValue(By locator, String value, String locatorName) throws Throwable {
		    boolean isByIndexSelected = false;
		    try {
		      WebElement webElement = getWebElement(locator);
		      Select dropDown = new Select(webElement);
		      dropDown.selectByValue(value);
		      Thread.sleep(5000);
		      isByIndexSelected = true;
		    } catch (Exception e) {
		      e.printStackTrace();
		    } finally {
		      if (isByIndexSelected) {
		    	  extentTest.log(LogStatus.PASS, " Option by value " + value
		            + " is selected from the drop down " + locatorName);
		      } else {
		    	  extentTest.log(LogStatus.FAIL,
		            "Option by value " + value + " is Not selected from the drop dpwn " + locatorName
		                + extentTest.addScreenCapture(getScreenshot(locatorName)));
		      }
		    }
		    return isByIndexSelected;
		  }
	  
	  
	  public boolean selectByVisibleText(By locator, String text, String locatorName) throws Throwable {
		    boolean isByIndexSelected = false;
		    try {
		      WebElement webElement = getWebElement(locator);
		      Select dropDown = new Select(webElement);
		      dropDown.selectByVisibleText(text);
		      Thread.sleep(5000);
		      isByIndexSelected = true;
		    } catch (Exception e) {
		      e.printStackTrace();
		    } finally {
		      if (isByIndexSelected) {
		    	  extentTest.log(LogStatus.PASS, " Option by Visible text " + text
		            + " is selected from the drop down " + locatorName);
		      } else {
		    	  extentTest.log(LogStatus.FAIL,
		            "Option by Visible text " + text + " is Not selected from the drop dpwn " + locatorName
		                + extentTest.addScreenCapture(getScreenshot(locatorName)));
		      }
		    }
		    return isByIndexSelected;
		  }
	  
	  
	  public void waitForElementToBeClickable(By locator, int withTime) throws Throwable {
		    try {
		      WebDriverWait wait = new WebDriverWait(driver, withTime);
		      wait.until(ExpectedConditions.elementToBeClickable(locator));
		    } catch (Throwable e) {
		      e.printStackTrace();
		    }
		  }
	  
	  public void fluentWaitForElementPresent(By locator, String locatorName, int withTimeout,
		      int pollingEvery) throws Throwable {
		    boolean flag = false;
		    try {
		      Wait<WebDriver> wait = new FluentWait<>(driver)
		          .withTimeout(Duration.ofSeconds(withTimeout))
		          .pollingEvery(Duration.ofSeconds(pollingEvery)).ignoring(WebDriverException.class);

		      wait.until(new Function<WebDriver, WebElement>() {
		        public WebElement apply(WebDriver driver) {
		          return getWebElement(locator);
		        }
		      });
		      flag = true;
		    } catch (Exception e) {
		      e.printStackTrace();
		    } finally {
		      if (flag) {
		    	  extentTest.log(LogStatus.PASS, "Successfully located element '" + locatorName + "'");
		      } else {
		    	  extentTest.log(LogStatus.FAIL, "Falied to locate element '" + locatorName + "'");
		      }
		    }
		  }
	  
	  public boolean acceptAlert() throws Throwable {
		    boolean isAlertAccepted = false;
		    try {
		      WebDriverWait wait = new WebDriverWait(driver, 5);
		      Alert alert = wait.until(ExpectedConditions.alertIsPresent());
		      alert.accept();
		      isAlertAccepted = true;
		    } catch (NoAlertPresentException ex) {
		      ex.printStackTrace();
		    } finally {
		      if (isAlertAccepted) {
		    	  extentTest.log(LogStatus.PASS, "Alet accepted");
		      } else {
		    	  extentTest.log(LogStatus.FAIL, "Failed to accept alert"
		            + extentTest.addScreenCapture(getScreenshot("Failed to accept alert!")));
		      }
		    }
		    return isAlertAccepted;
		  }
	  
	  public boolean dismissAlert() throws Throwable {
		    boolean isAlertAccepted = false;
		    try {
		      WebDriverWait wait = new WebDriverWait(driver, 5);
		      Alert alert = wait.until(ExpectedConditions.alertIsPresent());
		      alert.dismiss();
		      isAlertAccepted = true;
		    } catch (NoAlertPresentException ex) {
		      ex.printStackTrace();
		    } finally {
		      if (isAlertAccepted) {
		    	  extentTest.log(LogStatus.PASS, "Alet dismissed");
		      } else {
		    	  extentTest.log(LogStatus.FAIL, "Failed to dismiss alert"
		            + extentTest.addScreenCapture(getScreenshot("Failed to dismiss alert!")));
		      }
		    }
		    return isAlertAccepted;
		  }
	  
	  public boolean sendDataToAlert(String data) throws Throwable {
		    boolean isAlertAccepted = false;
		    try {
		      WebDriverWait wait = new WebDriverWait(driver, 3);
		      Alert alert = wait.until(ExpectedConditions.alertIsPresent());
		      alert.sendKeys(data);
		      isAlertAccepted = true;
		    } catch (NoAlertPresentException ex) {
		      ex.printStackTrace();
		    } finally {
		      if (isAlertAccepted) {
		    	  extentTest.log(LogStatus.PASS, "Alet handled with d input data");
		      } else {
		    	  extentTest.log(LogStatus.FAIL, "Failed to alert with data"
		            + extentTest.addScreenCapture(getScreenshot("Failed to alert with data!")));
		      }
		    }
		    return isAlertAccepted;
		  }
}
