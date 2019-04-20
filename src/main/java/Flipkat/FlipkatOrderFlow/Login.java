package Flipkat.FlipkatOrderFlow;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;

public class Login extends BrowserConfig{
	
	
	
	@FindBy(xpath=".//*[@class='LM6RPg']")
	private WebElement search;
	
	public void searchMobile()
	{
		BrowserConfig.launchURL("https://www.flipkart.com/");
		search.sendKeys("mobile");
		search.sendKeys(Keys.ENTER);
		driver.close();
		driver.quit();
	}

}
