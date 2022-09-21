package com.seetech.automation.pages;

import org.openqa.selenium.By;
import com.seetech.automation.actions.ActionEngine;

public class ToolsQAPage extends ActionEngine{
	private By elements = By.xpath("//h5[text()='Elements']");
	private By forms = By.xpath("//h5[text()='Forms']");
	
	public void clickElementsCard() throws Throwable {
		Thread.sleep(5000);
		click(elements, "Elements");
	}
	
	public void clickFormsCard() throws Throwable {
		click(forms, "Forms");
	}

}
