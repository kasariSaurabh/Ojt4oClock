package DriverFactory;

import org.openqa.selenium.WebDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import CommonFunLibrary.ERP_Functions;
import Utilities.ExcelFileUtil;

public class DataDrivenFramework {
WebDriver driver;
//reference object to call function
ERP_Functions erp=new ERP_Functions();
@BeforeTest
public void applaunch()throws Throwable
{
	String launch=erp.launchUrl("http://webapp.qedge.com");
	System.out.println(launch);
	String login=erp.verifyLogin("admin", "master");
	System.out.println(login);
}
@Test
public void stock() throws Throwable
{
	ExcelFileUtil xl=new ExcelFileUtil();
	int rc=xl.rowCount("Supplier");
	int cc=xl.colCount("Supplier");
	Reporter.log("no of row are::"+rc+" "+"no of columns are::"+cc,true);
	for(int i=1;i<=rc; i++)
	{
		String sname=xl.getCellData("Supplier", i, 0);
		String address=xl.getCellData("Supplier", i, 1);
		String city=xl.getCellData("Supplier", i, 2);
		String country=xl.getCellData("Supplier", i, 3);
		String cperson=xl.getCellData("Supplier", i, 4);
		String pnumber=xl.getCellData("Supplier", i, 5);
		String mail=xl.getCellData("Supplier", i, 6);
		String mnumber=xl.getCellData("Supplier", i, 7);
		String note=xl.getCellData("Supplier", i, 8);
		String results=erp.verifySupplier(sname, address, city, country, cperson, pnumber, mail, mnumber, note);
		xl.setCellData("Supplier", i, 9, results);
		
	}
}
@AfterTest
public void teardown()throws Throwable
{	
	String logoutapp=erp.verifyLogout();
	System.out.println(logoutapp);
	//driver.close();
}
}
