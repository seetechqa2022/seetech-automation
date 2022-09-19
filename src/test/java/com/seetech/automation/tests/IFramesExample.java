package com.seetech.automation.tests;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.testng.annotations.Test;

import com.seetech.automation.actions.ActionEngine;

public class IFramesExample extends ActionEngine{
	
	@Test
	public void frames() {
		driver.get("https://demoqa.com/frames");
		WebElement frame = driver.findElement(By.id("frame1"));
		driver.switchTo().frame(frame);
		
		String headerText = driver.findElement(By.id("sampleHeading")).getText();
		
		System.out.println("header test:: "+headerText);
		
		
		driver.switchTo().defaultContent();
		
		String pageHeaderText = driver.findElement(By.xpath("//*[@id=\"app\"]/div/div/div[1]/div")).getText();
		
		System.out.println("Page Header Text:: "+pageHeaderText);
	}
	

}
