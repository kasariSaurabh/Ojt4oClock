package CommonFunLibrary;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class ERP_Functions {

	WebDriver driver;

/*ProjctName:ERP_Stock
 * ModuleName:Launching url
 * TesterName:Saurabh
 */

public String launchUrl(String url) throws Throwable{
	String res="";
	System.setProperty("webdriver.chrome.driver", "./CommonJars/chromedriver.exe");
	driver = new ChromeDriver();
	driver.get(url);
	driver.manage().window().maximize();
	Thread.sleep(5000);
	if(driver.findElement(By.id("btnsubmit")).isDisplayed()){
		res = "Launch Application Success";
	}else{
		res = "Launch Application Fail";
		
	}
	return res;
	
}

/* ProjctName:ERP_Stock
 * ModuleName:Login
 * TesterName:Saurabh
 */
public String verifyLogin(String username, String password) throws Throwable
{
	String res="";
	WebElement user = driver.findElement(By.name("username"));
	user.clear();
	Thread.sleep(2000);
	user.sendKeys(username);
	Thread.sleep(2000);
	
	WebElement pass= driver.findElement(By.name("password"));
	pass.clear();
	Thread.sleep(2000);
	pass.sendKeys(password);
	Thread.sleep(2000);
	driver.findElement(By.name("btnsubmit")).click();
	Thread.sleep(2000);
	if(driver.findElement(By.id("logout")).isDisplayed()){
		res = "Login Success";
		
	}else{
		res = "Login Fail";
		
	}
	return res;
}

/* ProjctName:ERP_Stock
 * ModuleName:Logout
 * TesterName:Saurabh
 */

public String verifyLogout() throws Throwable{
	String res="";
	driver.findElement(By.xpath("//a[@id='logout']")).click();
	Thread.sleep(4000);
	driver.findElement(By.xpath("//button[@class='ajs-button btn btn-primary']")).click();
	if(driver.findElement(By.name("btnsubmit")).isDisplayed()){
		res="Application Logout Success";
	}else{
		res="Application Logout Fail";
	}
	Thread.sleep(4000);
	driver.close();
	return res;
	
}

/* ProjctName:ERP_Stock
 * ModuleName:Supplier Creation
 * TesterName:Saurabh
 */

public String verifySupplier(String sname, String address, String city, 
		String country, String cperson, String pnumber, String email, String mnumber, String note) throws Throwable{
	
	String res="";
	driver.findElement(By.linkText("Suppliers")).click();
	Thread.sleep(4000);
	driver.findElement(By.xpath("//div[@class='panel-heading ewGridUpperPanel']//span[@class='glyphicon glyphicon-plus ewIcon']")).click();
	Thread.sleep(3000);
	
	//get supplier number and store
	
	String Exp_data = driver.findElement(By.name("x_Supplier_Number")).getAttribute("value");
	driver.findElement(By.id("x_Supplier_Name")).sendKeys(sname);
	driver.findElement(By.id("x_Address")).sendKeys(address);
	driver.findElement(By.id("x_City")).sendKeys(city);
	driver.findElement(By.id("x_Country")).sendKeys(country);
	driver.findElement(By.id("x_Contact_Person")).sendKeys(cperson);
	driver.findElement(By.id("x_Phone_Number")).sendKeys(pnumber);
	driver.findElement(By.id("x__Email")).sendKeys(email);
	driver.findElement(By.id("x_Mobile_Number")).sendKeys(mnumber);
	driver.findElement(By.id("x_Notes")).sendKeys(note);
	
	driver.findElement(By.id("btnAction")).sendKeys(Keys.ENTER);
	Thread.sleep(4000);
	driver.findElement(By.xpath("//*[text()='OK!']")).click();
	Thread.sleep(4000);
	driver.findElement(By.xpath("(//*[text()='OK'])[6]")).click();
	Thread.sleep(4000);
	if(driver.findElement(By.id("psearch")).isDisplayed()){
		
	driver.findElement(By.id("psearch")).clear();
	Thread.sleep(4000);
	driver.findElement(By.id("psearch")).sendKeys(Exp_data);
	driver.findElement(By.id("btnsubmit")).click();
	Thread.sleep(4000);
	}else{
		
	driver.findElement(By.xpath("//span[@class='glyphicon glyphicon-search ewIcon']")).click();
	driver.findElement(By.id("psearch")).clear();
	Thread.sleep(4000);
	driver.findElement(By.id("psearch")).sendKeys(Exp_data);
	driver.findElement(By.id("btnsubmit")).click();
	Thread.sleep(4000);
	}
	
	String Act_Data = driver.findElement(By.xpath("//table[@id='tbl_a_supplierslist']/tbody/tr[1]/td[6]/div/span/span")).getText();
	Thread.sleep(4000);
	System.out.println(Exp_data+"       "+Act_Data);
	if(Exp_data.equals(Act_Data)){
		
		res = "Pass";
		
	}else{
		
		res = "Fail";
	}
	return res;
}
/*public static void main(String[] args) throws Throwable {
	
	ERP_Functions er = new ERP_Functions();
	er.launchUrl("http://webapp.qedge.com");
	er.verifyLogin("admin", "master");
	er.verifySupplier("Test", "Hyderabad", "Ameerpet", "INDIA", "Ramesh", "9857462525", "ramesh@125", "8541256854", "notes");
	er.verifyLogout();
}*/
}
