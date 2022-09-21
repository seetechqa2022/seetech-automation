package com.seetech.automation.tests;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.testng.annotations.Test;

import com.seetech.automation.actions.ActionEngine;
import com.seetech.automation.pages.ElementsPage;
import com.seetech.automation.pages.TextBoxPage;
import com.seetech.automation.pages.ToolsQAPage;

public class SeeTechTest extends ActionEngine{
	ToolsQAPage toolsQA = new ToolsQAPage();
	ElementsPage elementsPage = new ElementsPage();
	TextBoxPage textBoxPage = new TextBoxPage();
	
	//@Test
	public void test() throws Throwable {
		extentTest = extentReports.startTest("SriptName", "TestCase");
		toolsQA.clickElementsCard();
		elementsPage.verifyHeader();
		textBoxPage.clickTextBox();
		textBoxPage.fillTextBoxForm();
		Thread.sleep(5000);
	}
	
	//@Test
	public void test2() throws Throwable {
		driver.navigate().to("https://demoqa.com/automation-practice-form");
		WebElement we = driver.findElement(By.cssSelector("label[for='gender-radio-1']"));
		boolean flag = we.isSelected();
		
		System.out.println("Before Selection:: "+flag);
		if(flag == false) {
			we.click();
		}
		Thread.sleep(10000);
		flag = we.isSelected();
		System.out.println("After Selection:: "+flag);
		

		WebElement checkBox = driver.findElement(By.cssSelector("label[for='hobbies-checkbox-1']"));
		boolean flag2 = checkBox.isSelected();
		System.out.println("Before Check box Selection:: "+flag2);
		if(flag2 == false) {
			checkBox.click();
		}
		Thread.sleep(10000);
	}
	
	//@Test
	public void selectMenu() throws Throwable {
		extentTest = extentReports.startTest("SriptName", "TestCase");
		By selLocator = By.cssSelector("select[id='oldSelectMenu']");
		driver.navigate().to("https://demoqa.com/select-menu");
		//selectByIndex(selLocator, 10, "Select Drop down");
		//selectByValue(selLocator, "4", "Select Drop down");
		selectByVisibleText(selLocator, "White", "Select Drop down");
		waitForElementToBeClickable(selLocator, 10);
		click(selLocator, "some locator");
	}
	
	@Test
	public void fileUpload() throws Throwable {
		driver.navigate().to("https://demoqa.com/automation-practice-form");
		By fileUpload = By.id("uploadPicture");
		WebElement fwe = getWebElement(fileUpload);
		fwe.sendKeys("E:\\uma\\seetech-automation\\ConfigData\\config.properties");
		Thread.sleep(10000);
	}

}
