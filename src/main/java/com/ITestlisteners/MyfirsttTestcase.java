package com.ITestlisteners;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

@Listeners(com.ITestlisteners.TestNGListerner.class)
public class MyfirsttTestcase {
	
	@Test
	public void GoogleTitle() {
		WebDriverManager.chromedriver().setup();
		ChromeDriver driver = new ChromeDriver();
		//WebDriverManager.chromedriver().setup();
		driver.get("https://www.google.com/");
		driver.manage().window().maximize();
		System.out.println(driver.getTitle());
		String title = "Google";
		Assert.assertEquals(title, "Google");
		driver.close();
		
	}
	
	@Test
	public void Titlemach() {
		System.out.println("TestCase2 Dummy");
		Assert.assertTrue(true);
	

	}
}
