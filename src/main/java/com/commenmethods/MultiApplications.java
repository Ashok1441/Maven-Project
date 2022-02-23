package com.commenmethods;

import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;


//@Listeners(com.commenmethods.Browseropen.class)

public class MultiApplications {
	
	@Test
	//public static void main(String[] args) throws Throwable {
			public void sent() throws InterruptedException {
			WebDriverManager.chromedriver().setup();
			ChromeDriver driver = new ChromeDriver();
			driver.get("https://ratnaglobaltech.com/");
			//driver.manage().window().maximize();
			//driver.close();
			Thread.sleep(3000);
			
			}	//driver.get("https://www.flipkart.com/");
		

	  @Test
	   private void run() {
		  WebDriverManager.chromedriver().setup();
		  ChromeDriver driver= new ChromeDriver();
		  driver.get("https://www.flipkart.com/");
	  }
	 @Test
	 public void move() {
		 WebDriverManager.chromedriver().setup();
		 ChromeDriver driver = new ChromeDriver();
		 driver.get("https://mvnrepository.com/");

}
	 
}
