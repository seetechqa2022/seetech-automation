package com.seetech.automation.tests;

import org.testng.annotations.Test;

import com.seetech.automation.actions.ActionEngine;
import com.seetech.automation.pages.ElementsPage;
import com.seetech.automation.pages.TextBoxPage;
import com.seetech.automation.pages.ToolsQAPage;

public class SeeTechTest extends ActionEngine{
	ToolsQAPage toolsQA = new ToolsQAPage();
	ElementsPage elementsPage = new ElementsPage();
	TextBoxPage textBoxPage = new TextBoxPage();
	
	@Test
	public void test() throws Throwable {
		extentTest = extentReports.startTest("SriptName", "TestCase");
		toolsQA.clickElementsCard();
		elementsPage.verifyHeader();
		textBoxPage.clickTextBox();
		textBoxPage.fillTextBoxForm();
		Thread.sleep(5000);
	}

}
