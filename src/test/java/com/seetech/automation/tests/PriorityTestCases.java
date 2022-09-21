package com.seetech.automation.tests;

import org.testng.annotations.Test;

import com.seetech.automation.actions.ActionEngine;

public class PriorityTestCases extends ActionEngine {

	@Test(enabled = false)
	public void tekalign() {
		System.out.println("Tekalign");
	}
	
	@Test(priority = 1)
	public void tinsae() {
		System.out.println("Tinsae");
	}
	
	@Test(priority = 0)
	public void weldish() {
		System.out.println("Weldish");
	}

	@Test(priority = 3)
	public void yonas() {
		System.out.println("Yonas");
	}

	@Test(priority = 4)
	public void daniel() {
		System.out.println("Daniel");
	}

}
