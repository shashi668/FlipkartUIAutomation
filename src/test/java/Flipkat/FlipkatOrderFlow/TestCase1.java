package Flipkat.FlipkatOrderFlow;

import org.testng.annotations.Test;

public class TestCase1 {
	
	Login login = new Login();
	
	@Test
	public void login()
	{
		login.searchMobile();
		
	}

}
