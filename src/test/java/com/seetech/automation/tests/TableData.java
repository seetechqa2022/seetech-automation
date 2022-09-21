package com.seetech.automation.tests;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.testng.annotations.Test;

import com.seetech.automation.actions.ActionEngine;

public class TableData extends ActionEngine{
	
	@Test
	public void tableData() {
		driver.get("https://demo.guru99.com/test/table.html");
		WebElement table = driver.findElement(By.xpath("/html/body/table/tbody"));
		
		List<WebElement> rows = table.findElements(By.tagName("tr"));
		
		System.out.println(rows.size());
		
		
		for(int row=0; row<rows.size(); row++) {
			
			List<WebElement> cols = rows.get(row).findElements(By.tagName("td"));
			
			for(int col =0; col<cols.size(); col++) {
				WebElement data = cols.get(col);
				System.out.print(data.getText());
			}
			System.out.println();
		}
		
	}

}
