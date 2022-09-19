package com.seetech.automation.tests;

import java.util.Iterator;
import java.util.Set;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.testng.annotations.Test;

import com.seetech.automation.actions.ActionEngine;

public class WindowsHandling extends ActionEngine{
	
	
	@Test
	public void windows() {
		
		driver.get("https://demoqa.com/browser-windows");
		WebElement newWindow = driver.findElement(By.id("windowButton"));
		newWindow.click();
		
		
		String parentWindowId = driver.getWindowHandle();
		
		Set<String> windows = driver.getWindowHandles();
		
		System.out.println("Number of Windows opend:: "+windows.size());
		Iterator<String> itr = windows.iterator();
		while(itr.hasNext()) {
			String windowId= itr.next();
			System.out.println("Windo ID: "+windowId);
			if(!parentWindowId.equals(windowId)) {
				driver.switchTo().window(windowId);
				String header = driver.findElement(By.id("sampleHeading")).getText();
				System.out.println(header);
			}
			
		}
		
	}

}
