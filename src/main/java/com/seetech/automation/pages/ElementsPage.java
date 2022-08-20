package com.seetech.automation.pages;

import org.openqa.selenium.By;

import com.seetech.automation.actions.ActionEngine;

public class ElementsPage extends ActionEngine{
	private By elementsHeader =  By.xpath("//div[@class='main-header']");
	
	public void verifyHeader() throws Throwable{
		
		String headerText = getText(elementsHeader, "Elements Header");
		
		if("Elements".equals(headerText)) {
			System.out.println(" Header is verified successfully!!");
		}
		else {
			System.out.println(" Header verification failed!!");
		}
		
	}
}
