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
	

	@Test(priority = 7)
	public void tinsae2() {
		System.out.println("Tinsae");
	}

	@Test(priority = 6)
	public void teka() {
		System.out.println("Teka");

	}

	@Test(priority = 5)
	public void narasimha() {
		System.out.println("Narasimha");
	}
}
