package com.seetech.automation.pages;

import org.openqa.selenium.By;

import com.seetech.automation.actions.ActionEngine;

public class TextBoxPage extends ActionEngine {
	private By linkTextBox = By.xpath("//*[@id='item-0']/span[text()='Text Box']");
	private By textFullname = By.id("userName");
	private By textEmail = By.id("userEmail");
	private By buttonSubmit = By.id("submit");
	
	public void clickTextBox() throws Throwable {
		click(linkTextBox, "Link Text Box");
	}
	
	public void fillTextBoxForm() throws Throwable{
		type(textFullname, "Full Name", "Full Name!");
		type(textEmail, "emai@abc.com", "Email!");
		click(buttonSubmit, "Submit button!!");
	}
	
	

}
