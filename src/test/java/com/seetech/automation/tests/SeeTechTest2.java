package com.seetech.automation.tests;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.Test;

import com.seetech.automation.actions.ActionEngine;

public class SeeTechTest2 extends ActionEngine {
	
	//@Test
	public void alerts() throws Throwable{
		extentTest = extentReports.startTest("SriptName", "TestCase");
		driver.navigate().to("https://demoqa.com/alerts");
		
		click(By.id("promtButton"), "Click Button");
		Thread.sleep(10000);
		sendDataToAlert("Alert data");
		Thread.sleep(10000);
		
	}
	
	
	@Test
	public void actions() throws Throwable {
		/*Actions ac = new Actions(driver);
		ac.keyDown(Keys.CONTROL);
		ac.sendKeys("a");
		ac.keyUp(Keys.CONTROL);
		ac.perform();*/
		extentTest = extentReports.startTest("SriptName", "TestCase");
		driver.navigate().to("https://demoqa.com/text-box");
		
		/*
		 * new Actions(driver) .keyDown(Keys.CONTROL) .sendKeys("a")
		 * .keyUp(Keys.CONTROL) .perform();
		 */
		
		
//        WebElement textField = driver.findElement(By.id("userName"));
//        new Actions(driver)
//                .sendKeys(textField, "Selenium!")
//                .sendKeys(Keys.ARROW_LEFT)
//                .keyDown(Keys.SHIFT)
//                .sendKeys(Keys.ARROW_UP)
//                .keyUp(Keys.SHIFT)
//                .keyDown(Keys.CONTROL)
//                .sendKeys("xvv")
//                .keyUp(Keys.CONTROL)
//                .perform();
//        Thread.sleep(10000);
		
		Actions ac = new Actions(driver);
		WebElement textField = driver.findElement(By.id("userName"));
		textField.sendKeys("SeeTech");
		//ac.doubleClick(textField).perform();
		ac.contextClick(textField).perform();
		Thread.sleep(10000);
	}

}
