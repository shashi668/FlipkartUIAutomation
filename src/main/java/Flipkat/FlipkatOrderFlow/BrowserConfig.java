package Flipkat.FlipkatOrderFlow;

import java.io.File;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class BrowserConfig {
	
	public static WebDriver driver;
	
	public static WebDriver launchBrowser()
	{
		WebDriver driver = null;
		try
		{
			File file = new File("lib/chromedriver.exe");
			System.setProperty("webdriver.chrome.driver",file.getAbsolutePath());
			driver = new ChromeDriver();
			driver.manage().window().maximize();
		}catch (Exception e) {
			e.printStackTrace();
		}
		return driver;
	}
	
	public static WebDriver launchURL(String url)
	{
		try
		{
			driver = launchBrowser();
			driver.get(url);
		}catch (Exception e) {
			e.printStackTrace();
		}
		return driver;
	}

}
