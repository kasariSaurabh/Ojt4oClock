package DriverFactory;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import CommonFunLibrary.FunctionLibrary;
import Utilities.ExcelFileUtil;

public class DriverScript {
	WebDriver driver;
	ExtentReports report;
	ExtentTest test;
	//@Test
public void startTest() throws Throwable
{
	//create reference object for accessing excelutil methods
ExcelFileUtil excel=new ExcelFileUtil();
	//iterating all row in MasterTestCases sheet
for(int i=1; i<=excel.rowCount("MasterTestCases");i++)
{
	String ModuleStatus="";
if(excel.getCellData("MasterTestCases", i, 2).equalsIgnoreCase("Y"))
{
	
	//store module name into TCModule
String TCModule=excel.getCellData("MasterTestCases", i, 1);
//generate report under project
report=new ExtentReports("./Reports/"+" "+TCModule+" "+FunctionLibrary.generateDate()+ ".html");

//iterate all rows in TCModule
for(int j=1; j<=excel.rowCount(TCModule);j++)
{
	//test case excutes here
	test=report.startTest(TCModule);
	
	//read all columns in TCModule testcases
	String Description=excel.getCellData(TCModule, j, 0);
	String Object_Type=excel.getCellData(TCModule, j, 1);
	String Locator_Type=excel.getCellData(TCModule, j, 2);
	String Locator_Value=excel.getCellData(TCModule, j, 3);
	String Test_Data=excel.getCellData(TCModule, j, 4);
	try{
		if(Object_Type.equalsIgnoreCase("startBrowser"))
		{
			try{
			driver=FunctionLibrary.startBrowser(driver);
			test.log(LogStatus.INFO, Description);
			}finally{
				excel.setCellData("MasterTestCases", i, 3, "Fail");
			}
		}else if(Object_Type.equalsIgnoreCase("openApplication"))
		{
			try{
			FunctionLibrary.openApplication(driver);
			test.log(LogStatus.INFO, Description);
			}finally{
				excel.setCellData("MasterTestCases", i, 3, "Fail");
			}
		}else if(Object_Type.equalsIgnoreCase("waitForElement"))
		{
			try{
			FunctionLibrary.waitForElement(driver, Locator_Type, Locator_Value, Test_Data);
			test.log(LogStatus.INFO, Description);
			}finally{
				excel.setCellData("MasterTestCases", i, 3, "Fail");
			}
		}else if(Object_Type.equalsIgnoreCase("typeAction"))
		{
			try{
			FunctionLibrary.typeAction(driver, Locator_Type, Locator_Value, Test_Data);
			test.log(LogStatus.INFO, Description);
			}finally{
				excel.setCellData("MasterTestCases", i, 3, "Fail");
			}
		}else if(Object_Type.equalsIgnoreCase("clickAction"))
		{
			try{
			FunctionLibrary.clickAction(driver, Locator_Type, Locator_Value);
			test.log(LogStatus.INFO, Description);
			}finally{
				//ModuleStatus="false";
				excel.setCellData("MasterTestCases", i, 3, "Fail");
			}
			
		}else if(Object_Type.equalsIgnoreCase("closeBrowser"))
		{
			try{
			FunctionLibrary.closeBrowser(driver);
			test.log(LogStatus.INFO, Description);
			}finally{
				excel.setCellData("MasterTestCases", i, 3, "Fail");
			}
		}else if(Object_Type.equalsIgnoreCase("capturData"))
		{
			try{
			FunctionLibrary.capturData(driver, Locator_Type, Locator_Value);
			test.log(LogStatus.INFO, Description);
			}finally{
				excel.setCellData("MasterTestCases", i, 3, "Fail");
			}
		}else if(Object_Type.equalsIgnoreCase("tableValidation"))
		{
			try{
			FunctionLibrary.tableValidation(driver, Test_Data);
			test.log(LogStatus.INFO, Description);
			}finally{
				excel.setCellData("MasterTestCases", i, 3, "Fail");
			}
		}
		else if(Object_Type.equalsIgnoreCase("stockCategories"))
		{
			try{
			FunctionLibrary.stockCategories(driver);
			test.log(LogStatus.INFO, Description);
			}finally{
				excel.setCellData("MasterTestCases", i, 3, "Fail");
			}
		}else if(Object_Type.equalsIgnoreCase("stockValidation"))
		{
			try{
			FunctionLibrary.stockValidation(driver, Test_Data);
			test.log(LogStatus.INFO, Description);
			}finally{
				excel.setCellData("MasterTestCases", i, 3, "Fail");
			}
		}
		//write as pass into TCModule status column
		//excel.setCellData(TCModule, j, 5, "Pass");
		//test.log(LogStatus.PASS, Description);
		//ModuleStatus="true";
	}catch(Exception e){
		System.err.println(e.getMessage());
		//write as Fail into TCModule status column
		excel.setCellData(TCModule, j, 5, "Fail");
		test.log(LogStatus.FAIL, Description);
		
		//take screen shot into local system
		File screen=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
		//copy screen shot into local system
		FileUtils.copyFile(screen, new File("D:\\Ojt4oclock\\ERP_Accounting\\Screen\\"+Description+FunctionLibrary.generateDate()+".png"));
		//ModuleStatus="false";
		report.flush();
		report.endTest(test);
		break;
	}
	excel.setCellData(TCModule, j, 5, "Pass");
	test.log(LogStatus.PASS, Description);
	ModuleStatus="true";

	if(ModuleStatus.equalsIgnoreCase("True"))
	{
		excel.setCellData("MasterTestCases", i, 3, "Pass");
	}else if(ModuleStatus.equalsIgnoreCase("False"))
	{
		excel.setCellData("MasterTestCases", i, 3, "Fail");
	}
	
report.flush();
report.endTest(test);
}
}else{
	//write as not executed for testcases which are flag N
excel.setCellData("MasterTestCases", i, 3, "Not Executed");
}
}
}
}
