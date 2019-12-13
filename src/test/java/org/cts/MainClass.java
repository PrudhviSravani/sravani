package org.cts;

import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class MainClass extends BaseClass {
	public static void main(String[] args) throws IOException {
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\Dell\\Desktop\\Prudhvi\\ExcelIntegration\\driver\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.get("https://www.facebook.com/");
		WebElement user = driver.findElement(By.id("email"));
		user.sendKeys(getDataFromExcel(1,0));
		WebElement pass = driver.findElement(By.id("pass"));
		pass.sendKeys(getDataFromExcel(2,0));
	}

}
