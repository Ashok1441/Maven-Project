package com.ITestlisteners;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;


@Listeners(com.ITestlisteners.TestNGListerner.class)
public class TestNGbrowser {
	
	@Test
	public void Googleverify() {
		
		WebDriver driver = new ChromeDriver();
		driver.get("https://demo.actitime.com/");
		driver.manage().window().maximize();
		System.out.println(driver.getTitle());
		driver.close();
		
	}

}
