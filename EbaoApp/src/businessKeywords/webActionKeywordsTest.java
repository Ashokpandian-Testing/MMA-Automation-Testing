package businessKeywords;

import static org.junit.Assert.assertTrue;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.ss.util.SheetUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import allocator.CommonFunctions;
import allocator.Xml_Reporting_Functions;
import junit.framework.Assert;
import jxl.WorkbookSettings;
import jxl.write.WritableWorkbook;
import allocator.Caller;

public class webActionKeywordsTest {

	public static CommonFunctions CF = new CommonFunctions();
	public static Xml_Reporting_Functions XF = new Xml_Reporting_Functions();
	public static String testCaseID = "";
	public static String testFlowSheetName = "";
	public static WebDriver driver =  null;
	public static String objectName= "";
	public static String objectNameGiven= "";
	public static String action= "";
	public static String actionRowNo= "";
	public static String inputData;
	public static String dataSheetValue= "";
	public static String browser= "";
	public static String objectProperty = "";
	public static String objectType = "";
	public static int objectSheetRowCount = 1;
	public static int reportSheetRowNo = 1;
	public static int customizedReportSheetRowNo = 1;
	public static int customizedReportSheetColCount = 1;
	public static int dataSheetRowCount = 1;
	public static int dataSheetTestCaseRowIndex = -1;
	public static String takeSnapShot= "";
	public static String BusinessComponents = "";
	 public static String ReportObjectName;
	 public static String BSDFinalResultPath;
	 public static String BSDScreenshotPath;
	 public static String BSDFinalScreenPath,BSDExecutionDashboardPath,BSDFinalDashboardPath,POSSheetName;;
	 public static String BSDSummaryReportPath,BSDFinalReportPath;
	 public static    String parenthandle;

	public static Map<String, String> map = new HashMap<String, String>();
	public static int AutomationStepCount;
	public static int UtilitySheetCount;
	public static String PageName;



	public static void main(String[] args) 
	{
		// TODO Auto-generated method stub


	}


	// Method for sleep time
	public void sleep() throws InterruptedException
	{
		try
		{
			int secTime = Integer.parseInt(inputData);
			for(int i=1; i<=secTime; i++)
			{
				Thread.sleep(1000l);
			}
			
		
		
		}catch (Exception e)
		{
	
			 
	
		}

	}
	
	public  void SwitchFrame() {
		try
		{
		  driver.switchTo().frame(inputData);
		 Caller.ExecutionStatus="PASS";	      
		 Caller.Reason=webActionKeywordsTest.action+"  action is Executed SuccessFully ";
		 Caller.stringBuilder.append(Caller.Reason);	
		 Caller.stringBuilder.append("\n");
		 
		 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
		 Caller.StringExecutionStatus.append("\n");
		
		}catch (Exception e)
		{
			e.getMessage();
			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=webActionKeywordsTest.action+" action is NOT Executed SuccessFully ";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			 
	
		}

	}
	public   void  close_driver(){
		try {
   
     	 driver.close();
       driver.switchTo().window(parenthandle);
       Caller.ExecutionStatus="PASS";	      
		 Caller.Reason=webActionKeywordsTest.action+" is closed SuccessFully ";
		 Caller.stringBuilder.append(Caller.Reason);	
		 Caller.stringBuilder.append("\n");
		 
		 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
		 Caller.StringExecutionStatus.append("\n");
       
		}catch(Exception e) {
			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=webActionKeywordsTest.action+"is NOT closed SuccessFully ";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
	    	  
	      }
        }

	

	public   WebDriver  getHandleToWindow(){
        WebDriver popup = null;
      try {
    	  
        parenthandle=driver.getWindowHandle();
        Set<String> windowIterator = driver.getWindowHandles();

        for (String s : windowIterator) 
        {
          String windowHandle = s; 
     	 popup = driver.switchTo().window(windowHandle);
          if (popup.getTitle().matches(inputData))
          {
        
        	 return popup;
         
          }
          Caller.ExecutionStatus="PASS";	      
 		 Caller.Reason=webActionKeywordsTest.action+"  is handled SuccessFully ";
 		 Caller.stringBuilder.append(Caller.Reason);	
 		 Caller.stringBuilder.append("\n");
 		 
 		 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
 		 Caller.StringExecutionStatus.append("\n");
            
        } 
      }catch(Exception e) {
    		 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=webActionKeywordsTest.action+" NOT handled SuccessFully ";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
    	  
      }
	return popup;
        }
	
	

	public   WebDriver  getHandleToWindowWithoutParentWindow(){
        WebDriver popup = null;
      try {
    	  
      
        Set<String> windowIterator = driver.getWindowHandles();

        for (String s : windowIterator) 
        {
          String windowHandle = s; 
     	 popup = driver.switchTo().window(windowHandle);
          if (popup.getTitle().matches(inputData))
          {
        
        	 return popup;
          }
          Caller.ExecutionStatus="PASS";	      
  		 Caller.Reason=webActionKeywordsTest.action+"  is handled SuccessFully ";
  		 Caller.stringBuilder.append(Caller.Reason);	
  		 Caller.stringBuilder.append("\n");
  		 
  		 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
  		 Caller.StringExecutionStatus.append("\n");
             
            
        } 
      }catch(Exception e) {
    	  Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=webActionKeywordsTest.action+" is NOT handled SuccessFully ";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
 	  
    	  
      }
	return popup;
        }
  
  
	// Method for closing browser
	public  void  closeBrowser() throws InterruptedException, IOException
	{

		try 
		{
			browser = browser.toLowerCase();
			if(browser.startsWith("firefox"))
			{
				Runtime.getRuntime().exec("taskkill /F /IM firefox.exe");
				
				Thread.sleep(500);
			}else if(browser.startsWith("ie"))
			{
				Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe");
				
				Thread.sleep(500);
			}else if(browser.startsWith("chrome"))
			{
				Runtime.getRuntime().exec("taskkill /F /IM chrome.exe");
				
				Thread.sleep(500);
			}

			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason="The "+browser+" is closed Successfully";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(Caller.Reason, "1");
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");

		} 

		
		 catch (Exception e)
		{
			e.getMessage();
			//System.out.println(e);
			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason="The"+browser+"is not  closed Successfully";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(Caller.Reason, "3");
			 XF.XmlscreenShot(false);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			 
		}
	}

	public  void SelectBox() throws InterruptedException {
		try 
		{
		if (objectType.equalsIgnoreCase("css")) {
		
		WebElement mySelectRel = driver.findElement(By.cssSelector(objectProperty));
		Select dropdownele= new Select(mySelectRel);
		dropdownele.selectByVisibleText(inputData);
		}
		else	if (objectType.equalsIgnoreCase("xpath")) {
			
			WebElement mySelectRel = driver.findElement(By.xpath(objectProperty));
			Select dropdownele= new Select(mySelectRel);
			dropdownele.selectByVisibleText(inputData);
			}
		else	if (objectType.equalsIgnoreCase("name")) {
			
			WebElement mySelectRel = driver.findElement(By.name(objectProperty));
			Select dropdownele= new Select(mySelectRel);
			dropdownele.selectByVisibleText(inputData);
			}
		else	if (objectType.equalsIgnoreCase("id")) {
			
			WebElement mySelectRel = driver.findElement(By.id(objectProperty));
			Select dropdownele= new Select(mySelectRel);
			dropdownele.selectByVisibleText(inputData);
			}
		webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
		if(webActionKeywordsTest.PageName.contains("/")) {
			webActionKeywordsTest.PageName.replace("/","_");
		}
		 Caller.ExecutionStatus="PASS";	      
		 Caller.Reason="The '"+inputData+"' has been selected SuccessFully on '"+webActionKeywordsTest.ReportObjectName+"' "+"field of '"+webActionKeywordsTest.PageName+"'";
		 Caller.stringBuilder.append(Caller.Reason);	
		 Caller.stringBuilder.append("\n");
		 XF.AddResult(Caller.Reason, "1");
		 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
		 Caller.StringExecutionStatus.append("\n");
		}
		catch (Exception e)
		{
			e.getMessage();
			
			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=e.toString();
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(Caller.Reason, "3");
			 XF.XmlscreenShot(false);			
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			 
		}
	

		
	}
	// Method for set text
	public  void  setText() throws InterruptedException
	{
		try 
		{
			if (webActionKeywordsTest.inputData.contains("Map")) {
			String mapKeyName = inputData;
			String mapDataStored = map.get(mapKeyName).toString();
			inputData=mapDataStored;
			}
			if(objectType.equalsIgnoreCase("css"))
			{
				driver.findElement(By.cssSelector(objectProperty)).clear();
				driver.findElement(By.cssSelector(objectProperty)).sendKeys(inputData);
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				driver.findElement(By.xpath(objectProperty)).clear();
				driver.findElement(By.xpath(objectProperty)).sendKeys(inputData);
			}else if(objectType.equalsIgnoreCase("name"))
			{
				driver.findElement(By.name(objectProperty)).clear();
				driver.findElement(By.name(objectProperty)).sendKeys(inputData);
			}else if(objectType.equalsIgnoreCase("id"))
			{
				driver.findElement(By.id(objectProperty)).clear();
				driver.findElement(By.id(objectProperty)).sendKeys(inputData);
			}
			webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
			if(webActionKeywordsTest.PageName.contains("/")) {
				webActionKeywordsTest.PageName.replace("/","_");
			}
			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason="The '"+inputData+"' has been entered SuccessFully on '"+webActionKeywordsTest.ReportObjectName+"' "+"field of '"+webActionKeywordsTest.PageName+"'";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(Caller.Reason, "1");
//			 XF.XmlscreenShot(true);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");

		}catch (Exception e)
		{
			e.getMessage();

			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=e.toString();
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(Caller.Reason, "3");
			 XF.XmlscreenShot(false);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			 
		}
	}
	
	
	// Method forVerifyTextReadOnly
	public  void  VerifyTextReadOnly() throws InterruptedException
	{
		String ReadOnlyValue = null;
		try 
		{
			if(objectType.equalsIgnoreCase("css"))
			{
				ReadOnlyValue=	driver.findElement(By.cssSelector(objectProperty)).getAttribute("readonly") ;
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				ReadOnlyValue=	driver.findElement(By.xpath(objectProperty)).getAttribute("readonly");
			}else if(objectType.equalsIgnoreCase("name"))
			{
				ReadOnlyValue=driver.findElement(By.name(objectProperty)).getAttribute("readonly");
			}else if(objectType.equalsIgnoreCase("id"))
			{
				ReadOnlyValue=driver.findElement(By.id(objectProperty)).getAttribute("readonly");
			}
			webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
			if(webActionKeywordsTest.PageName.contains("/")) {
				webActionKeywordsTest.PageName.replace("/","_");
			}
			if(ReadOnlyValue.equalsIgnoreCase("True")) {

			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason="The '"+webActionKeywordsTest.ReportObjectName+"' "+"is in ReadOnly field in the page '"+webActionKeywordsTest.PageName+"'";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(Caller.Reason, "1");
//			 XF.XmlscreenShot(true);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			}
			else {

				 Caller.ExecutionStatus="FAIL";	      
				 Caller.Reason="The '"+webActionKeywordsTest.ReportObjectName+"' "+"is in NOT ReadOnly field in the page '"+webActionKeywordsTest.PageName+"'";
				 Caller.stringBuilder.append(Caller.Reason);	
				 Caller.stringBuilder.append("\n");
				 XF.AddResult(Caller.Reason, "3");
				 XF.XmlscreenShot(false);
				 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
				 Caller.StringExecutionStatus.append("\n");
				
			}

		}catch (Exception e)
		{
			e.getMessage();

			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=e.toString();
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(Caller.Reason, "3");
			 XF.XmlscreenShot(false);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			 

		}
	}
	
	public static void clickRadioinTable()   {
		 
		 
		 
	    try    {

			String objectNameArr[] = objectName.split(",");

			//Getting the Web Table property
			CF.getMappedObjectPropertyFromObjectSheet(objectNameArr[0],webActionKeywordsTest.PageName);
			CF.getObjectProperty_Type(objectName);
			
			String objectPropertyTbl = objectProperty;

			//Getting the page Next Button Property
			CF.getMappedObjectPropertyFromObjectSheet(objectNameArr[1],webActionKeywordsTest.PageName);
			CF.getObjectProperty_Type(objectName);
			
			String objectPropertyNextBtn = objectProperty;
	    	
		
		   	String ColumnObject=objectPropertyNextBtn+"[1]/td";
		    WebElement Webtable= webActionKeywordsTest.driver.findElement(By.xpath(objectPropertyTbl)); 
		    java.util.List<WebElement>TotalRowCount=Webtable.findElements(By.xpath(objectPropertyNextBtn));
	
		    java.util.List<WebElement>TotalColCount=Webtable.findElements(By.xpath(ColumnObject));
		   
		 
		    String first_part = objectPropertyNextBtn+"[";
		    String second_part = "]/td[";
		    String third_part = "]";
		    Loop:
		    for (int i=1; i<=TotalRowCount.size(); i++){
		    	  //Used for loop for number of columns.
		    	  for(int j=1; j<=TotalColCount.size(); j++){
		    	   //Prepared final xpath of specific cell as per values of i and j.
		    	   String final_xpath = first_part+i+second_part+j+third_part;
		    	   //Will retrieve value from located cell and print It.
		    	   String Table_data = webActionKeywordsTest.driver.findElement(By.xpath(final_xpath)).getText();
		    	   if(Table_data.contains(inputData)) { 
		    		
		    		   List  oRadioButton = webActionKeywordsTest.driver.findElements(By.name("addressradio"));
		    		 int Value=i-2;
		    		   
		    		   ((WebElement) oRadioButton.get(Value)).click();
		    		   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
		   			if(webActionKeywordsTest.PageName.contains("/")) {
		   				webActionKeywordsTest.PageName.replace("/","_");
		   			}
		   			Caller.ExecutionStatus="PASS";	      
					 Caller.Reason="' "+webActionKeywordsTest.action+"' clicked successfully in the page '"+webActionKeywordsTest.PageName+"'";
					 Caller.stringBuilder.append(Caller.Reason);	
					 Caller.stringBuilder.append("\n");
					 XF.AddResult(Caller.Reason, "1");
//					 XF.XmlscreenShot(true);
					 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
					 Caller.StringExecutionStatus.append("\n");
		    	   break Loop;	
		    		
		    	   
	                 }
		    		 
	           }
	           }
	    }
	          catch(Exception e){
	        	  
	        		e.getMessage();

	   			 Caller.ExecutionStatus="FAIL";	      
	   			 Caller.Reason=e.toString();
	   			 Caller.stringBuilder.append(Caller.Reason);	
	   			 Caller.stringBuilder.append("\n");
	   			 XF.AddResult(Caller.Reason, "3");
				 XF.XmlscreenShot(false);
	   			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
	   			 Caller.StringExecutionStatus.append("\n");
	 	 
	    }
	}

	

	
	// Method forVerifyTextReadOnly
		public  void  ClickButtonifExist() throws InterruptedException
		{
	WebElement Button;
			try 
			{
				if(objectType.equalsIgnoreCase("css"))
				{
					Button=driver.findElement(By.cssSelector(objectProperty));
					if(Button.isDisplayed()) {
						Button.click();	
					}
				}else if(objectType.equalsIgnoreCase("xpath"))
				{
					Button=driver.findElement(By.xpath(objectProperty));
					if(Button.isDisplayed()) {
						Button.click();	
					}
					
				}else if(objectType.equalsIgnoreCase("name"))
				{
					Button=driver.findElement(By.name(objectProperty));
					if(Button.isDisplayed()) {
						Button.click();	
					}
					
				}else if(objectType.equalsIgnoreCase("id"))
				{
				Button=driver.findElement(By.id(objectProperty));
				if(Button.isDisplayed()) {
					Button.click();	
				}
				}
				   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
		   			if(webActionKeywordsTest.PageName.contains("/")) {
		   				webActionKeywordsTest.PageName.replace("/","_");
		   			}

				 Caller.ExecutionStatus="PASS";	      
				 Caller.Reason="The button '"+webActionKeywordsTest.ReportObjectName+"' "+"is  displayed in the page '"+webActionKeywordsTest.PageName+"' "+"it is clicked";
				 Caller.stringBuilder.append(Caller.Reason);	
				 Caller.stringBuilder.append("\n");
				 XF.AddResult(Caller.Reason, "1");
//				 XF.XmlscreenShot(true);
				 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
				 Caller.StringExecutionStatus.append("\n");
		}
			
			
			catch (Exception e)
			{
				//e.getMessage();

				 Caller.ExecutionStatus="PASS";	      
				 Caller.Reason="The button'"+webActionKeywordsTest.ReportObjectName+"'"+"field  is  displayed in the page'"+webActionKeywordsTest.PageName+"'"+"it is not clicked";
				 Caller.stringBuilder.append(Caller.Reason);	
				 Caller.stringBuilder.append("\n");
				 XF.AddResult(Caller.Reason, "3");
				 XF.XmlscreenShot(false);
				 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
				 Caller.StringExecutionStatus.append("\n");
				 
			}
		}
		public static void keyPress()   {
			try {
				WebElement Button;
			
			if(objectType.equalsIgnoreCase("css"))
				
			{
				Button=driver.findElement(By.cssSelector(objectProperty));
				if(Button.isDisplayed()) {
					Button.sendKeys(Keys.ENTER);	
				}
				
			}else if(objectType.equalsIgnoreCase("link"))
			{
				Button=driver.findElement(By.linkText(objectProperty));
				if(Button.isDisplayed()) {
					Button.sendKeys(Keys.ENTER);	
				}
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				Button=driver.findElement(By.xpath(objectProperty));
				if(Button.isDisplayed()) {
					Button.sendKeys(Keys.ENTER);	
				}
			}else if(objectType.equalsIgnoreCase("id"))
			{
				Button=driver.findElement(By.id(objectProperty));
				if(Button.isDisplayed()) {
					Button.sendKeys(Keys.ENTER);	
				}
			}
				else if(objectType.equalsIgnoreCase("name"))
			{
				Button=driver.findElement(By.name(objectProperty));
				if(Button.isDisplayed()) {
					Button.sendKeys(Keys.ENTER);	
				}
			}
			else if(objectType.equalsIgnoreCase("class"))
			{
				Button=driver.findElement(By.className(objectProperty));
				if(Button.isDisplayed()) {
					Button.sendKeys(Keys.ENTER);	
				}
			}
			   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
	   			if(webActionKeywordsTest.PageName.contains("/")) {
	   				webActionKeywordsTest.PageName.replace("/","_");
	   			}

			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason="Enter key was not pressed in the pop up successfully";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(Caller.Reason, "1");
//			 XF.XmlscreenShot(true);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
	
			}
			catch (Exception e)
			{
				 Caller.ExecutionStatus="FAIL";	      
				 Caller.Reason="Enter key was not pressed in the pop up successfully";
				 Caller.stringBuilder.append(Caller.Reason);	
				 Caller.stringBuilder.append("\n");
				 XF.AddResult(Caller.Reason, "3");
				 XF.XmlscreenShot(false);
				 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
				 Caller.StringExecutionStatus.append("\n");
			}
		
		}
		
		public static void clickinTable()   {
			 
			 
			 
		    try    {

				String objectNameArr[] = objectName.split(",");

				//Getting the Web Table property
				CF.getMappedObjectPropertyFromObjectSheet(objectNameArr[0],webActionKeywordsTest.PageName);
				CF.getObjectProperty_Type(objectName);
				
				String objectPropertyTbl = objectProperty;

				//Getting the page Next Button Property
				CF.getMappedObjectPropertyFromObjectSheet(objectNameArr[1],webActionKeywordsTest.PageName);
				CF.getObjectProperty_Type(objectName);
				
				String objectPropertyNextBtn = objectProperty;
		    	
			
			   	String ColumnObject=objectPropertyNextBtn+"[1]/td";
			    WebElement Webtable= webActionKeywordsTest.driver.findElement(By.xpath(objectPropertyTbl)); 
			    java.util.List<WebElement>TotalRowCount=Webtable.findElements(By.xpath(objectPropertyNextBtn));
		
			    java.util.List<WebElement>TotalColCount=Webtable.findElements(By.xpath(ColumnObject));
			   
			 
			    String first_part = objectPropertyNextBtn+"[";
			    String second_part = "]/td[";
			    String third_part = "]";
			    Loop:
			    for (int i=2; i<=TotalRowCount.size(); i++){
			    	  //Used for loop for number of columns.
			    	  for(int j=2; j<=TotalColCount.size(); j++){
			    	   //Prepared final xpath of specific cell as per values of i and j.
			    	   String final_xpath = first_part+i+second_part+j+third_part;
			    	   //Will retrieve value from located cell and print It.
			    	   String Table_data = webActionKeywordsTest.driver.findElement(By.xpath(final_xpath)).getText();
			    	   if(Table_data.equalsIgnoreCase(inputData)) { 
			    		
			    		   WebElement  oRadioButton = webActionKeywordsTest.driver.findElement(By.xpath(final_xpath));
			    		   
			    		   Caller.ExecutionStatus="PASS";	      
				    	   Caller.Reason= "'Policy Status is : "+inputData+"'as expected'";
							 Caller.stringBuilder.append(Caller.Reason);	
							 Caller.stringBuilder.append("\n");
							 XF.AddResult(Caller.Reason, "1");
//							 XF.XmlscreenShot(true);
							 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
							 Caller.StringExecutionStatus.append("\n");
			    		   
			    	  		Actions action =new Actions(driver);
			    	  		action.doubleClick(oRadioButton).perform();
			    	  	   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
				   			if(webActionKeywordsTest.PageName.contains("/")) {
				   				webActionKeywordsTest.PageName.replace("/","_");
				   			}
			    	  	   Caller.ExecutionStatus="PASS";	      
				    	   Caller.Reason="'"+webActionKeywordsTest.action+"'clicked successfully in the page"+webActionKeywordsTest.PageName+"'";
							 Caller.stringBuilder.append(Caller.Reason);	
							 Caller.stringBuilder.append("\n");
							 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
							 Caller.StringExecutionStatus.append("\n");
			    		 
			    	   break Loop;	                     
		                 }
			    	
			    	  
		           }
		           }
		    }
		          catch(Exception e){
		        		 
		        	  Caller.ExecutionStatus="FAIL";	      
			    	   Caller.Reason="'"+webActionKeywordsTest.action+"'  NOT clicked successfully in the page"+webActionKeywordsTest.PageName+"'";
						 Caller.stringBuilder.append(Caller.Reason);	
						 Caller.stringBuilder.append("\n");
						 XF.AddResult(Caller.Reason, "3");
						 XF.XmlscreenShot(false);
						 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
						 Caller.StringExecutionStatus.append("\n");
		    }
		}

		public static void keyPressRunTime()   {
			try {
				WebElement Button;
			
			if(objectType.equalsIgnoreCase("css"))
				
			{
				Button=driver.findElement(By.cssSelector(objectProperty));
				if(Button.isDisplayed()) {
					Button.sendKeys(inputData);		
				}
				
			}else if(objectType.equalsIgnoreCase("link"))
			{
				Button=driver.findElement(By.linkText(objectProperty));
				if(Button.isDisplayed()) {
					Button.sendKeys(inputData);	
				}
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				Button=driver.findElement(By.xpath(objectProperty));
				if(Button.isDisplayed()) {
					Button.sendKeys(inputData);	
				}
			}else if(objectType.equalsIgnoreCase("id"))
			{
				Button=driver.findElement(By.id(objectProperty));
				if(Button.isDisplayed()) {
					Button.sendKeys(inputData);	
				}
			}
				else if(objectType.equalsIgnoreCase("name"))
			{
				Button=driver.findElement(By.name(objectProperty));
				if(Button.isDisplayed()) {
					Button.sendKeys(inputData);		
				}
			}
			else if(objectType.equalsIgnoreCase("class"))
			{
				Button=driver.findElement(By.className(objectProperty));
				if(Button.isDisplayed()) {
					Button.sendKeys(inputData);		
				}
			}
			   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
	   			if(webActionKeywordsTest.PageName.contains("/")) {
	   				webActionKeywordsTest.PageName.replace("/","_");
	   			}

			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason=webActionKeywordsTest.action+"is Executed Successfully";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
	
			}
			catch (Exception e)
			{
				 Caller.ExecutionStatus="PASS";	      
				 Caller.Reason=webActionKeywordsTest.action+"is Executed Successfully";
				 Caller.stringBuilder.append(Caller.Reason);	
				 Caller.stringBuilder.append("\n");
				 XF.AddResult(e.toString(), "3");
				 XF.XmlscreenShot(false);
				 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
				 Caller.StringExecutionStatus.append("\n");
			}
		
		}
		
	       
        // Method for get text
        public  String  getTextbyAttribute() throws InterruptedException
  
        {
               try 
               {
                      
                      String mapKeyName = inputData.split(",")[0].toString();
                      String mapDataStored = map.get(mapKeyName).toString();
                      String Attributetype=inputData.split(",")[1].toString();
                      String getText= "";

                      if(objectType.equalsIgnoreCase("css"))
                      {
                             getText = driver.findElement(By.cssSelector(objectProperty)).getAttribute(Attributetype).trim();
                      }else if(objectType.equalsIgnoreCase("xpath"))
                      {
                             getText = driver.findElement(By.xpath(objectProperty)).getAttribute(Attributetype).trim();
                      }else if(objectType.equalsIgnoreCase("id"))
                      {
                             getText = driver.findElement(By.id(objectProperty)).getAttribute(Attributetype).trim();
                      }else if(objectType.equalsIgnoreCase("name"))
                      {
                             getText = driver.findElement(By.name(objectProperty)).getAttribute(Attributetype).trim();
                      }else if(objectType.equalsIgnoreCase("link"))
                      {
                             getText = driver.findElement(By.linkText(objectProperty)).getAttribute(Attributetype).trim();
                      }
                         webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
                            if(webActionKeywordsTest.PageName.contains("/")) {
                                   webActionKeywordsTest.PageName.replace("/","_");
                            }
                      //map.put(inputData.split("=")[1], getText);
                      map.put(mapDataStored, getText);
                      Caller.ExecutionStatus="PASS";        
                       Caller.Reason="The"+webActionKeywordsTest.ReportObjectName+"'"+"value is retrived successfully in the page"+webActionKeywordsTest.PageName+"'";
                      Caller.stringBuilder.append(Caller.Reason);  
                      Caller.stringBuilder.append("\n");
                      
                       Caller.StringExecutionStatus.append(Caller.ExecutionStatus);      
                      Caller.StringExecutionStatus.append("\n");
        

                      return getText;

               }catch (Exception e)
               {
                      e.getMessage();
                      Caller.ExecutionStatus="FAIL";        
                       Caller.Reason="The"+webActionKeywordsTest.ReportObjectName+"'"+"value is NOT retrived successfully in the page"+webActionKeywordsTest.PageName+"'";
                      Caller.stringBuilder.append(Caller.Reason);  
                      Caller.stringBuilder.append("\n");
                      
                       Caller.StringExecutionStatus.append(Caller.ExecutionStatus);      
                      Caller.StringExecutionStatus.append("\n");
                      
               }
               return "";
        }


	// Method forVerifyTextReadOnly
		public  void  VerifyButtonEnabled() throws InterruptedException
		{
			 boolean Flag = false;
            
			try 
			{
				
				if(objectType.equalsIgnoreCase("css"))
				{
					Flag=driver.findElement(By.cssSelector(objectProperty)).isEnabled();

				}else if(objectType.equalsIgnoreCase("xpath"))
				{
				Flag=driver.findElement(By.xpath(objectProperty)).isEnabled();
				}else if(objectType.equalsIgnoreCase("name"))
				{
				Flag=driver.findElement(By.name(objectProperty)).isEnabled();
				}else if(objectType.equalsIgnoreCase("id"))
				{
				Flag=driver.findElement(By.id(objectProperty)).isEnabled();
			

			}
				   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
		   			if(webActionKeywordsTest.PageName.contains("/")) {
		   				webActionKeywordsTest.PageName.replace("/","_");
		   			}
				if (Flag==true) {

					 Caller.ExecutionStatus="PASS";	      
					 Caller.Reason="The '"+webActionKeywordsTest.ReportObjectName+"' "+"Object is Enabled in the page'"+webActionKeywordsTest.PageName+"'";
					 Caller.stringBuilder.append(Caller.Reason);	
					 Caller.stringBuilder.append("\n");
					 XF.AddResult(Caller.Reason, "1");
//					 XF.XmlscreenShot(true);
					 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
					 Caller.StringExecutionStatus.append("\n");
					
				}
				else {
					 Caller.ExecutionStatus="FAIL";	      
					 Caller.Reason="The '"+webActionKeywordsTest.ReportObjectName+"' "+"Object is Disabled in the page '"+webActionKeywordsTest.PageName+"'";
					 Caller.stringBuilder.append(Caller.Reason);	
					 Caller.stringBuilder.append("\n");
					 XF.AddResult(Caller.Reason, "3");
					 XF.XmlscreenShot(false);
					 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
					 Caller.StringExecutionStatus.append("\n");
						
				}
			}catch (Exception e)
			{
				e.getMessage();

				 Caller.ExecutionStatus="FAIL";	      
				 Caller.Reason=e.toString();
				 Caller.stringBuilder.append(Caller.Reason);	
				 Caller.stringBuilder.append("\n");
				 XF.AddResult(e.toString(), "3");
				 XF.XmlscreenShot(false);
				 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
				 Caller.StringExecutionStatus.append("\n");
				 

			}
		}
		
        public static void keyPressParentHandle()   {
            try {
                  WebElement Button;
            
            if(objectType.equalsIgnoreCase("css"))
                  
            {
                  Button=driver.findElement(By.cssSelector(objectProperty));
                  if(Button.isDisplayed()) {
                         Button.sendKeys(Keys.ENTER);     
                  }
                  
            }else if(objectType.equalsIgnoreCase("link"))
            {
                  Button=driver.findElement(By.linkText(objectProperty));
                  if(Button.isDisplayed()) {
                         Button.sendKeys(Keys.ENTER);     
                  }
            }else if(objectType.equalsIgnoreCase("xpath"))
            {
                  Button=driver.findElement(By.xpath(objectProperty));
                  if(Button.isDisplayed()) {
                         Button.sendKeys(Keys.ENTER);     
                  }
            }else if(objectType.equalsIgnoreCase("id"))
            {
                  Button=driver.findElement(By.id(objectProperty));
                  if(Button.isDisplayed()) {
                         Button.sendKeys(Keys.ENTER);     
                  }
            }
                  else if(objectType.equalsIgnoreCase("name"))
            {
                  Button=driver.findElement(By.name(objectProperty));
                  if(Button.isDisplayed()) {
                         Button.sendKeys(Keys.ENTER);     
                  }
            }
            else if(objectType.equalsIgnoreCase("class"))
            {
                  Button=driver.findElement(By.className(objectProperty));
                  if(Button.isDisplayed()) {
                         Button.sendKeys(Keys.ENTER);     
                  }
            }
               
         driver.switchTo().window(parenthandle);
            Caller.ExecutionStatus="PASS";        
             Caller.Reason=webActionKeywordsTest.action+"is Executed Successfully";
            Caller.stringBuilder.append(Caller.Reason);  
            Caller.stringBuilder.append("\n");
            XF.AddResult(Caller.Reason, "1");
             Caller.StringExecutionStatus.append(Caller.ExecutionStatus);     
            Caller.StringExecutionStatus.append("\n");

            }
            catch (Exception e)
            {
                  Caller.ExecutionStatus="PASS";        
                   Caller.Reason=webActionKeywordsTest.action+"is Executed Successfully";
                  Caller.stringBuilder.append(Caller.Reason);  
                  Caller.stringBuilder.append("\n");
                  XF.AddResult(e.toString(), "3");
 				 XF.XmlscreenShot(false);
                   Caller.StringExecutionStatus.append(Caller.ExecutionStatus);     
                  Caller.StringExecutionStatus.append("\n");
            }
     
     }

		
		public  String  getHiddenText() throws InterruptedException
        {
               try 
               {
                      String script= "";
                      JavascriptExecutor js = (JavascriptExecutor)webActionKeywordsTest.driver;
                      //String content = (String) ((JavascriptExecutor)        webActionKeywordsTest.driver).executeScript("return arguments[0].value", element);
                      //JavascriptExecutor je = (JavascriptExecutor)  webActionKeywordsTest.driver;
               
               
 if(objectType.equalsIgnoreCase("id"))
                      {
                             script = (String) js.executeScript("return document.getElementById('"+objectProperty+"').value;");
                      }else if(objectType.equalsIgnoreCase("name"))
                      {
                             script = (String) js.executeScript("return document.getElementsByName('"+objectProperty+"')[0].value;");
                      }
                         webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
                            if(webActionKeywordsTest.PageName.contains("/")) {
                                   webActionKeywordsTest.PageName.replace("/","_");
                            }
                      //map.put(inputData.split("=")[1], getText);
                      map.put(inputData, script);
                      Caller.ExecutionStatus="PASS";        
                       Caller.Reason="The"+webActionKeywordsTest.ReportObjectName+"'"+"value is retrived successfully in the page"+webActionKeywordsTest.PageName+"'";
                      Caller.stringBuilder.append(Caller.Reason);  
                      Caller.stringBuilder.append("\n");
                      
                        Caller.StringExecutionStatus.append(Caller.ExecutionStatus);      
                      Caller.StringExecutionStatus.append("\n");
        

                      return script;

               }catch (Exception e)
               {
                      e.getMessage();
                      Caller.ExecutionStatus="FAIL";        
                       Caller.Reason="The"+webActionKeywordsTest.ReportObjectName+"'"+"value is NOT retrived successfully in the page"+webActionKeywordsTest.PageName+"'";
                      Caller.stringBuilder.append(Caller.Reason);  
                      Caller.stringBuilder.append("\n");
                      
                       Caller.StringExecutionStatus.append(Caller.ExecutionStatus);      
                      Caller.StringExecutionStatus.append("\n");
                      
               }
               finally
          { 
              System.out.println("Executes whether exception occurs or not"); 
          } 
               return "";
        }

		public  String  getAttributeTitleVal() throws InterruptedException
        {
               try 
               {
                     String getText= "";

                     if(objectType.equalsIgnoreCase("css"))
                     {
                            getText = driver.findElement(By.cssSelector(objectProperty)).getAttribute("title").trim();
                     }else if(objectType.equalsIgnoreCase("xpath"))
                     {
                            getText = driver.findElement(By.xpath(objectProperty)).getAttribute("title").trim();
                     }else if(objectType.equalsIgnoreCase("id"))
                     {
                            getText = driver.findElement(By.id(objectProperty)).getAttribute("title").trim();
                     }else if(objectType.equalsIgnoreCase("name"))
                     {
                            getText = driver.findElement(By.name(objectProperty)).getAttribute("title").trim();
                     }else if(objectType.equalsIgnoreCase("link"))
                     {
                            getText = driver.findElement(By.linkText(objectProperty)).getAttribute("title").trim();
                     }
                        webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
                            if(webActionKeywordsTest.PageName.contains("/")) {
                                   webActionKeywordsTest.PageName.replace("/","_");
                            }
                     //map.put(inputData.split("=")[1], getText);
                     map.put(inputData, getText);
                     Caller.ExecutionStatus="PASS";        
                      Caller.Reason="The"+webActionKeywordsTest.ReportObjectName+"'"+"value is retrived successfully in the page"+webActionKeywordsTest.PageName+"'";
                     Caller.stringBuilder.append(Caller.Reason);  
                     Caller.stringBuilder.append("\n");
                     
                      Caller.StringExecutionStatus.append(Caller.ExecutionStatus);     
                     Caller.StringExecutionStatus.append("\n");
        

                     return getText;

               }catch (Exception e)
               {
                     e.getMessage();
                     Caller.ExecutionStatus="FAIL";        
                      Caller.Reason="The"+webActionKeywordsTest.ReportObjectName+"'"+"value is NOT retrived successfully in the page"+webActionKeywordsTest.PageName+"'";
                     Caller.stringBuilder.append(Caller.Reason);  
                     Caller.stringBuilder.append("\n");
                     
                      Caller.StringExecutionStatus.append(Caller.ExecutionStatus);     
                     Caller.StringExecutionStatus.append("\n");
                     
               }
               return "";
        }

	
		// Method forVerifyTextReadOnly
				public  void  VerifyButtonDisabled() throws InterruptedException
				{
					 String Flag = "";
		            
					try 
					{
						
						if(objectType.equalsIgnoreCase("css"))
						{
							Flag=driver.findElement(By.cssSelector(objectProperty)).getAttribute("disabled");

						}else if(objectType.equalsIgnoreCase("xpath"))
						{
						Flag=driver.findElement(By.xpath(objectProperty)).getAttribute("disabled");
						}else if(objectType.equalsIgnoreCase("name"))
						{
						Flag=driver.findElement(By.name(objectProperty)).getAttribute("disabled");
						}else if(objectType.equalsIgnoreCase("id"))
						{
						Flag=driver.findElement(By.id(objectProperty)).getAttribute("disabled");
					

					}
						   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
				   			if(webActionKeywordsTest.PageName.contains("/")) {
				   				webActionKeywordsTest.PageName.replace("/","_");
				   			}
						if (Flag.equalsIgnoreCase("true")) {

							 Caller.ExecutionStatus="PASS";	      
							 Caller.Reason="The '"+webActionKeywordsTest.ReportObjectName+"' "+"Object is Disabled in the page '"+webActionKeywordsTest.PageName+"'";
							 Caller.stringBuilder.append(Caller.Reason);	
							 Caller.stringBuilder.append("\n");
							 XF.AddResult(Caller.Reason, "1");
//							 XF.XmlscreenShot(true);
							 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
							 Caller.StringExecutionStatus.append("\n");
							
						}
						else {
							 Caller.ExecutionStatus="FAIL";	      
							 Caller.Reason="The '"+webActionKeywordsTest.ReportObjectName+"' "+"Object is Enabled in the page'"+webActionKeywordsTest.PageName+"'";
							 Caller.stringBuilder.append(Caller.Reason);	
							 Caller.stringBuilder.append("\n");
							 XF.AddResult(Caller.Reason, "3");
							 XF.XmlscreenShot(false);
							 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
							 Caller.StringExecutionStatus.append("\n");
								
						}
					}catch (Exception e)
					{
						e.getMessage();

						 Caller.ExecutionStatus="FAIL";	      
						 Caller.Reason=e.toString();
						 Caller.stringBuilder.append(Caller.Reason);	
						 Caller.stringBuilder.append("\n");
						 XF.AddResult(e.toString(), "3");
						 XF.XmlscreenShot(false);
						 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
						 Caller.StringExecutionStatus.append("\n");
						 
					}
				}

				public static void RegExMatches() {
		              try {
		                     
		              
		                     Boolean dataFetchFlag = false;
		                     String mapKeyName = inputData.split(",")[0].toString();
		                     String mapDataStored ;
		                     String expectedData = inputData.split(",")[1];
		                     if (expectedData.contains("capture")){
		                           mapDataStored = map.get(mapKeyName).split("\n")[0].toString();
		                     }
		                     else {
		                           mapDataStored = map.get(mapKeyName).split("\n")[1].toString();
		                     }
		                     if(!expectedData.contains("'"))
		                     {
		                           dataFetchFlag = CF.getDataFromDataSheet(expectedData,"DataSheet");
		                           if(dataFetchFlag == true)
		                           {
		                                  expectedData = dataSheetValue;
		                           }else
		                           {
		                                  try
		                                  {
		                                         expectedData = map.get(expectedData).toString();
		                                  }catch (Exception e)
		                                  {
		                                         expectedData = "";
		                                  }
		                           }
		                     }else if(expectedData.contains("'"))
		                     {
		                           expectedData = expectedData.replace("'", "");
		                     }
		                        webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
		                           if(webActionKeywordsTest.PageName.contains("/")) {
		                                  webActionKeywordsTest.PageName.replace("/","_");
		                           }
		                           Pattern p = Pattern.compile(expectedData);
		                           Matcher m = p.matcher(mapDataStored);  
		                           boolean b = m.matches();
		                           if(b==true) {
		                                  Caller.ExecutionStatus="PASS";        
		                                   Caller.Reason="ActualValue '"+mapDataStored+"' is matched with ExpectedValue '"+expectedData+"'";             
		                                   Caller.stringBuilder.append(Caller.Reason);  
		                                  Caller.stringBuilder.append("\n");
		                                  XF.AddResult(Caller.Reason, "1");
		                     			 XF.XmlscreenShot(true);
		                                   Caller.StringExecutionStatus.append(Caller.ExecutionStatus);       
		                                  Caller.StringExecutionStatus.append("\n");
		                           }
		                           else {
		                                  Caller.ExecutionStatus="FAIL";         
		                                   Caller.Reason="ActualValue '"+mapDataStored+"' is NOT matched with ExpectedValue '"+expectedData+"'";             
		                                   Caller.stringBuilder.append(Caller.Reason);  
		                                  Caller.stringBuilder.append("\n");
		                                  XF.AddResult(Caller.Reason, "3");
			                     			 XF.XmlscreenShot(false);
		                                   Caller.StringExecutionStatus.append(Caller.ExecutionStatus);       
		                                  Caller.StringExecutionStatus.append("\n");
		                                  
		                            }
		                     
		              }
		              catch(Exception e) {
		                     e.printStackTrace();
		                     Caller.ExecutionStatus="FAIL";        
		                      Caller.Reason=e.toString();
		                     Caller.stringBuilder.append(Caller.Reason);  
		                     Caller.stringBuilder.append("\n");
		                     XF.AddResult(e.toString(), "3");
                 			 XF.XmlscreenShot(false);
		                      Caller.StringExecutionStatus.append(Caller.ExecutionStatus);     
		                     Caller.StringExecutionStatus.append("\n");
		                     
		              }
		       }



		public static void DaysCalculation() throws Exception{
		       try {
		       Boolean dataFetchFlag = false;
		       String mapKeyName = inputData.split(",")[0].toString();
		       String mapDataStored = map.get(mapKeyName).toString();
		       String expectedData = inputData.split(",")[1];
		       if(!expectedData.contains("'"))
		       {
		              dataFetchFlag = CF.getDataFromDataSheet(expectedData,"DataSheet");
		              if(dataFetchFlag == true)
		              {
		                     expectedData = dataSheetValue;
		              }else
		              {
		                     try
		                     {
		                           expectedData = map.get(expectedData).toString();
		                     }catch (Exception e)
		                     {
		                           expectedData = "";
		                     }
		              }
		       }else if(expectedData.contains("'"))
		       {
		              expectedData = expectedData.replace("'", "");
		       }
		       SimpleDateFormat sdfmt1 = new SimpleDateFormat("MM/dd/yyyy");
		  Date date1=sdfmt1.parse(mapDataStored);
		Date date2=sdfmt1.parse(expectedData);
		          long diff=date1.getTime()-date2.getTime();
		          float days=diff/(1000*60*60*24);
		          String noofdays=Float.toString(days).substring(0, 3);
		             map.put(mapKeyName, noofdays);
		       }
		       catch(Exception e) {
		              e.printStackTrace();
		       }
		       }


	// Method for get text
	public  String  getText() throws InterruptedException
	{
		try 
		{
			String getText= "";

			if(objectType.equalsIgnoreCase("css"))
			{
				getText = driver.findElement(By.cssSelector(objectProperty)).getAttribute("value").trim();
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				getText = driver.findElement(By.xpath(objectProperty)).getAttribute("value").trim();
			}else if(objectType.equalsIgnoreCase("id"))
			{
				getText = driver.findElement(By.id(objectProperty)).getAttribute("value").trim();
			}else if(objectType.equalsIgnoreCase("name"))
			{
				getText = driver.findElement(By.name(objectProperty)).getAttribute("value").trim();
			}else if(objectType.equalsIgnoreCase("link"))
			{
				getText = driver.findElement(By.linkText(objectProperty)).getAttribute("value").trim();
			}
			   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
	   			if(webActionKeywordsTest.PageName.contains("/")) {
	   				webActionKeywordsTest.PageName.replace("/","_");
	   			}
			//map.put(inputData.split("=")[1], getText);
			map.put(inputData, getText);
			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason="The"+webActionKeywordsTest.ReportObjectName+"'"+"value is retrived successfully in the page"+webActionKeywordsTest.PageName+"'";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
	

			return getText;

		}catch (Exception e)
		{
			e.getMessage();
			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason="The"+webActionKeywordsTest.ReportObjectName+"'"+"value is NOT retrived successfully in the page"+webActionKeywordsTest.PageName+"'";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(e.toString(), "3");
			 XF.XmlscreenShot(false);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			 
		}
		return "";
	}

	
	public  String  getValue() throws InterruptedException
	{
		try 
		{
			String getText= "";

			if(objectType.equalsIgnoreCase("css"))
			{
				getText = driver.findElement(By.cssSelector(objectProperty)).getText();
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				getText = driver.findElement(By.xpath(objectProperty)).getText();
			}else if(objectType.equalsIgnoreCase("id"))
			{
				getText = driver.findElement(By.id(objectProperty)).getText();
			}else if(objectType.equalsIgnoreCase("name"))
			{
				getText = driver.findElement(By.name(objectProperty)).getText();
			}else if(objectType.equalsIgnoreCase("link"))
			{
				getText = driver.findElement(By.linkText(objectProperty)).getText();
			}
			   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
	   			if(webActionKeywordsTest.PageName.contains("/")) {
	   				webActionKeywordsTest.PageName.replace("/","_");
	   			}

			//map.put(inputData.split("=")[1], getText);
			map.put(inputData, getText);

			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason=webActionKeywordsTest.action+" with the object '"+webActionKeywordsTest.ReportObjectName+"'"+"value is retrived successfully in the page"+webActionKeywordsTest.PageName+"'";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
	

			return getText;

		}catch (Exception e)
		{
			e.getMessage();
			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=webActionKeywordsTest.action+" with the object"+webActionKeywordsTest.ReportObjectName+"'"+"value is NOT retrived successfully in the page"+webActionKeywordsTest.PageName+"'";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(e.toString(), "3");
			 XF.XmlscreenShot(false);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			 
		}
		return "";
	}



	// Method for Verifying Screen Text
	public  void  compareText() throws InterruptedException
	{
		try 
		{
			Boolean dataFetchFlag = false;
			String mapKeyName = inputData.split(",")[0].toString();
			
			String mapDataStored = map.get(mapKeyName).toString();
			if (mapDataStored.equalsIgnoreCase("Y")){
				mapDataStored="Yes";
			}
			if (mapDataStored.equalsIgnoreCase("N")){
				mapDataStored="No";
			} 
			String expectedData = inputData.split(",")[1];
			if(!expectedData.contains("'"))
			{
				dataFetchFlag = CF.getDataFromDataSheet(expectedData,"DataSheet");
				if(dataFetchFlag == true)
				{
					expectedData = dataSheetValue;
				}else
				{
					try
					{
						expectedData = map.get(expectedData).toString();
					}catch (Exception e)
					{
						expectedData = "";
					}
				}
			}else if(expectedData.contains("'"))
			{
				expectedData = expectedData.replace("'", "");
			}
			   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
	   			if(webActionKeywordsTest.PageName.contains("/")) {
	   				webActionKeywordsTest.PageName.replace("/","_");
	   			}

			if(mapDataStored.equalsIgnoreCase(expectedData))
			{
				 Caller.ExecutionStatus="PASS";	      
				 Caller.Reason="Actual Value '"+mapDataStored+"' is matched with ExpectedValue '"+expectedData+"'";             
				 Caller.stringBuilder.append(Caller.Reason);	
				 Caller.stringBuilder.append("\n");
				 XF.AddResult(Caller.Reason, "1");
				 XF.XmlscreenShot(true);
				 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
				 Caller.StringExecutionStatus.append("\n");

			}else
			{
				 Caller.ExecutionStatus="FAIL";	      
				 Caller.Reason="ActualValue '"+mapDataStored+"' is NOT matched with ExpectedValue '"+expectedData+"'";             
				 Caller.stringBuilder.append(Caller.Reason);	
				 Caller.stringBuilder.append("\n");
				 XF.AddResult(Caller.Reason, "3");
				 XF.XmlscreenShot(false);
				 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
				 Caller.StringExecutionStatus.append("\n");

		
			}

		}catch (Exception e)
		{
			e.getMessage();
			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=e.toString();           
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(e.toString(), "3");
			 XF.XmlscreenShot(false);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
	
			 
		}
	}




	// Method for Verifying Screen Text
	public  void  getTextCompare() throws InterruptedException
	{
		try 
		{
			String getText= "";

			if(objectType.equalsIgnoreCase("css"))
			{
				getText = driver.findElement(By.cssSelector(objectProperty)).getText().trim();
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				getText = driver.findElement(By.xpath(objectProperty)).getText().trim();
			}else if(objectType.equalsIgnoreCase("id"))
			{
				getText = driver.findElement(By.id(objectProperty)).getText().trim();
			}else if(objectType.equalsIgnoreCase("name"))
			{
				getText = driver.findElement(By.name(objectProperty)).getText().trim();
			}else if(objectType.equalsIgnoreCase("link"))
			{
				getText = driver.findElement(By.linkText(objectProperty)).getText().trim();
			}
			   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
	   			if(webActionKeywordsTest.PageName.contains("/")) {
	   				webActionKeywordsTest.PageName.replace("/","_");
	   			}
			//map.put(inputData.split("=")[1], getText);
			map.put(inputData.split(",")[0].toString(), getText);



			Boolean dataFetchFlag = false;
			String mapKeyName = inputData.split(",")[0].toString();
			String mapDataStored = map.get(mapKeyName).toString();
			String expectedData = inputData.split(",")[1];
			if(!expectedData.contains("'"))
			{
				dataFetchFlag = CF.getDataFromDataSheet(expectedData,"DataSheet");
				if(dataFetchFlag == true)
				{
					expectedData = dataSheetValue;
				}else
				{
					try
					{
						expectedData = map.get(expectedData).toString();
					}catch (Exception e)
					{
						expectedData = "";
					}
				}
			}else if(expectedData.contains("'"))
			{
				expectedData = expectedData.replace("'", "");
			}

			if(mapDataStored.equalsIgnoreCase(expectedData))
			{
				//System.out.println(mapDataStored + " and "+ expectedData + " matched. Status PASS.");
				 Caller.ExecutionStatus="PASS";	      
				 Caller.Reason="ActualValue '"+mapDataStored+"' is matched with ExpectedValue '"+expectedData+"'";             
				 Caller.stringBuilder.append(Caller.Reason);	
				 Caller.stringBuilder.append("\n");
				 XF.AddResult(Caller.Reason, "1");
				 XF.XmlscreenShot(true);
				 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
				 Caller.StringExecutionStatus.append("\n");
				
			}else
			{
				
				 Caller.ExecutionStatus="FAIL";	      
				 Caller.Reason="ActualValue '"+mapDataStored+"' is NOT matched with ExpectedValue '"+expectedData+"'";             
				 Caller.stringBuilder.append(Caller.Reason);	
				 Caller.stringBuilder.append("\n");
				 XF.AddResult(Caller.Reason, "3");
				 XF.XmlscreenShot(false);
				 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
				 Caller.StringExecutionStatus.append("\n");
			}

		}catch (Exception e)
		{
			e.getMessage();
			 
			 XF.AddResult(e.toString(), "3");
			 XF.XmlscreenShot(false);
		}
	}




	// Method for click
	public  void  click() throws InterruptedException
	{
		try 
		{
			if(objectType.equalsIgnoreCase("css"))
			{
				driver.findElement(By.cssSelector(objectProperty)).click();
			}else if(objectType.equalsIgnoreCase("link"))
			{
				driver.findElement(By.linkText(objectProperty)).click();
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				driver.findElement(By.xpath(objectProperty)).click();
			}else if(objectType.equalsIgnoreCase("id"))
			{
				driver.findElement(By.id(objectProperty)).click();
			}else if(objectType.equalsIgnoreCase("name"))
			{
				driver.findElement(By.name(objectProperty)).click();
			}
			else if(objectType.equalsIgnoreCase("class"))
			{
				driver.findElement(By.className(objectProperty)).click();
			}
			   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
	   			if(webActionKeywordsTest.PageName.contains("/")) {
	   				webActionKeywordsTest.PageName.replace("/","_");
	   			}
			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason="'"+webActionKeywordsTest.ReportObjectName+"' "+"has been clicked on the page '"+webActionKeywordsTest.PageName+"' "+"Successfully";             
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(Caller.Reason, "1");
//			 XF.XmlscreenShot(true);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");

		}
		
		catch (Exception e)
		{
			e.getMessage();

			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=e.toString();
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(e.toString(), "3");
			 XF.XmlscreenShot(false);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			 
		}
		}
	



	// Method for Clear Text Box
	public  void  clear() throws InterruptedException
	{
		try 
		{
			if(objectType.equalsIgnoreCase("css"))
			{
				driver.findElement(By.cssSelector(objectProperty)).clear();
			}else if(objectType.equalsIgnoreCase("link"))
			{
				driver.findElement(By.linkText(objectProperty)).clear();
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				driver.findElement(By.xpath(objectProperty)).clear();
			}else if(objectType.equalsIgnoreCase("id"))
			{
				driver.findElement(By.id(objectProperty)).clear();
			}else if(objectType.equalsIgnoreCase("name"))
			{
				driver.findElement(By.name(objectProperty)).clear();
			}

	
			reportSheetRowNo = reportSheetRowNo +1;

		}catch (Exception e)
		{
			e.getMessage();
		
			 XF.AddResult(e.toString(), "3");
			 XF.XmlscreenShot(false);
		}
	}


	// Method for refresh page
	public  void  refresh() throws InterruptedException
	{
		try 
		{
			driver.navigate().refresh();

	
			reportSheetRowNo = reportSheetRowNo +1;

		}catch (Exception e)
		{
			e.getMessage();
			//System.out.println(e);
		
			 XF.AddResult(e.toString(), "3");
			 XF.XmlscreenShot(false);
			
		}
	}




	// Method for going back page of browser
	public  void  back() throws InterruptedException
	{
		try 
		{
			driver.navigate().back();

	

		}catch (Exception e)
		{
			e.getMessage();
			//System.out.println(e);
			
			 XF.AddResult(e.toString(), "3");
			 XF.XmlscreenShot(false);
		}
	}



	// Method for going forward page of browser
	public  void  forward() throws InterruptedException
	{
		try 
		{
			driver.navigate().forward();


		}catch (Exception e)
		{
			e.getMessage();
			//System.out.println(e);
	
			
			
		}
	}



	// Method for going to new URL in browser
	public  void  gotoURL() throws InterruptedException
	{
		try 
		{
			driver.get(inputData);


		}catch (Exception e)
		{
			e.getMessage();
			//System.out.println(e);

			
			
		}
	}


	// Method for getting the URL of the Browser
	public  void  getURL() throws InterruptedException
	{
		try 
		{
			String currentURL = driver.getCurrentUrl();
			map.put(inputData, currentURL);

		

		}catch (Exception e)
		{
			e.getMessage();
			//System.out.println(e);
		
			
			
		}
	}




	// Method for getting the URL of the Browser
	public  void  getBrowserTitle() throws InterruptedException
	{
		try 
		{
			String browserTitle = driver.getTitle();
			map.put(inputData, browserTitle);

		

		}catch (Exception e)
		{
			e.getMessage();
			//System.out.println(e);
		
			
			
		}
	}

	// Method for Select from Drop Down List
	public  void  selectListItem() throws InterruptedException
	{
		try 
		{
			if(objectType.equalsIgnoreCase("css"))
			{
				new Select(driver.findElement(By.cssSelector(objectProperty))).selectByVisibleText(inputData);
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				new Select(driver.findElement(By.xpath(objectProperty))).selectByVisibleText(inputData);
			}else if(objectType.equalsIgnoreCase("name"))
			{
				new Select(driver.findElement(By.name(objectProperty))).selectByVisibleText(inputData);
			}else if(objectType.equalsIgnoreCase("id"))
			{
				new Select(driver.findElement(By.id(objectProperty))).selectByVisibleText(inputData);
			}else if(objectType.equalsIgnoreCase("link"))
			{
				new Select(driver.findElement(By.linkText(objectProperty))).selectByVisibleText(inputData);
			}
			   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
	   			if(webActionKeywordsTest.PageName.contains("/")) {
	   				webActionKeywordsTest.PageName.replace("/","_");
	   			}
			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason="The '"+inputData+"' has been selected SuccessFully on '"+webActionKeywordsTest.ReportObjectName+"' "+"field of '"+webActionKeywordsTest.PageName+"'";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(Caller.Reason, "1");
//			 XF.XmlscreenShot(true);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");

		}catch (Exception e)
		{
			e.getMessage();
			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=e.toString();
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(e.toString(), "3");
			 XF.XmlscreenShot(false);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			 
				
		}
	}


	 public  void DoubleClick() throws InterruptedException {
		 try {
		if (objectType.equalsIgnoreCase("css"))
			{
			WebElement Element= driver.findElement(By.cssSelector(objectProperty));
	  		Actions action =new Actions(driver);
	  		action.doubleClick(Element).perform();
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				WebElement Element= driver.findElement(By.xpath(objectProperty));
		  		Actions action =new Actions(driver);
		  		action.doubleClick(Element).perform();
			}else if(objectType.equalsIgnoreCase("name"))
			{
				WebElement Element= driver.findElement(By.name(objectProperty));
		  		Actions action =new Actions(driver);
		  		action.doubleClick(Element).perform();
			}else if(objectType.equalsIgnoreCase("id"))
			{
				WebElement Element= driver.findElement(By.id(objectProperty));
		  		Actions action =new Actions(driver);
		  		action.doubleClick(Element).perform();
			}else if(objectType.equalsIgnoreCase("link"))
			{
				WebElement Element= driver.findElement(By.linkText(objectProperty));
		  		Actions action =new Actions(driver);
		  		action.doubleClick(Element).perform();
			}
		   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
  			if(webActionKeywordsTest.PageName.contains("/")) {
  				webActionKeywordsTest.PageName.replace("/","_");
  			}
		 Caller.ExecutionStatus="PASS";	      
		 Caller.Reason="'"+webActionKeywordsTest.ReportObjectName+"' "+"has been Double Clicked on the page '"+webActionKeywordsTest.PageName+"' "+"Successfully";             
		 Caller.stringBuilder.append(Caller.Reason);	
		 Caller.stringBuilder.append("\n");
		 XF.AddResult(Caller.Reason, "1");
//		 XF.XmlscreenShot(true);
		 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
		 Caller.StringExecutionStatus.append("\n");

		 }
		 catch(Exception e) {

			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=e.toString();
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(e.toString(), "3");
			 XF.XmlscreenShot(false);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			 
				
			 
		 }
	  		
	  	
	      	
	      }

	// Method for maximize
	public  void  maximize() throws InterruptedException
	{
		try 
		{
			driver.manage().window().maximize();
			   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
	   			if(webActionKeywordsTest.PageName.contains("/")) {
	   				webActionKeywordsTest.PageName.replace("/","_");
	   			}
			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason="The maximize Action is Executed SuccessFully for the PageName'"+webActionKeywordsTest.PageName+"'"+" and the ObjectName is' "+webActionKeywordsTest.ReportObjectName+"'";	              
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");

		}
		catch (Exception e)
		{
			e.getMessage();
			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=e.toString();
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			 
				
		}
	}
	

	// Method for verifying element in screen
	public void verifyElementOnScreen() {
		try 
		{
			Boolean elementFound = false;
			
			if(objectType.equalsIgnoreCase("css"))
			{
				elementFound = CF.isElementPresent(By.cssSelector(objectProperty));
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				elementFound = CF.isElementPresent(By.xpath(objectProperty));
			}else if(objectType.equalsIgnoreCase("name"))
			{
				elementFound = CF.isElementPresent(By.name(objectProperty));
			}else if(objectType.equalsIgnoreCase("id"))
			{
				elementFound = CF.isElementPresent(By.id(objectProperty));
			}else if(objectType.equalsIgnoreCase("link"))
			{
				elementFound = CF.isElementPresent(By.linkText(objectProperty));
			}
			   webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
	   			if(webActionKeywordsTest.PageName.contains("/")) {
	   				webActionKeywordsTest.PageName.replace("/","_");
	   			}

			if(elementFound == true)
			{
			
				 Caller.ExecutionStatus="PASS";	      
				 Caller.Reason="'"+webActionKeywordsTest.ReportObjectName+"' "+"has been Verified on the page '"+webActionKeywordsTest.PageName+"' "+"Successfully";             
				 Caller.stringBuilder.append(Caller.Reason);	
				 Caller.stringBuilder.append("\n");
				 XF.AddResult(Caller.Reason, "1");
				 XF.XmlscreenShot(true);
				 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
				 Caller.StringExecutionStatus.append("\n");
		
			}else
			{
			
				 Caller.ExecutionStatus="PASS";	       
				 Caller.Reason="'"+webActionKeywordsTest.ReportObjectName+"' "+"has been Verified on the page '"+webActionKeywordsTest.PageName+"' "+"Successfully";             
				 Caller.stringBuilder.append(Caller.Reason);	
				 Caller.stringBuilder.append("\n");
				 XF.AddResult(Caller.Reason, "3");
				 XF.XmlscreenShot(false);
				 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
				 Caller.StringExecutionStatus.append("\n");
				 
					
			}



		} catch (Exception e)
		{
			e.getMessage();
			
			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=e.toString();
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(e.toString(), "3");
			 XF.XmlscreenShot(false);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			 
				
		}
	}


	// Method for verifying page name in screen
			public void verifyPageNameOnScreen() {
				try 
				{
					
					if(driver.getTitle().matches(inputData))
					{
						
						 Caller.ExecutionStatus="PASS";	      
						 Caller.Reason="The Screen '"+driver.getTitle()+"' has been Displayed Successfully";             
						 Caller.stringBuilder.append(Caller.Reason);	
						 Caller.stringBuilder.append("\n");
						 XF.AddResult(Caller.Reason, "1");
//						 XF.XmlscreenShot(true);
						 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
						 Caller.StringExecutionStatus.append("\n");
				
					}
					else
					{
						
						 Caller.ExecutionStatus="FAIL";	      
						 Caller.Reason="The Screen '"+driver.getTitle()+"' has not been Displayed Successfully";                   
						 Caller.stringBuilder.append(Caller.Reason);	
						 Caller.stringBuilder.append("\n");
						 XF.AddResult(Caller.Reason, "3");
						 XF.XmlscreenShot(false);
						 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
						 Caller.StringExecutionStatus.append("\n");
						 
							
					}
					

				} catch (Exception e)
				{
					e.getMessage();
					
					 Caller.ExecutionStatus="FAIL";	      
					 Caller.Reason=e.toString();
					 Caller.stringBuilder.append(Caller.Reason);	
					 Caller.stringBuilder.append("\n");		
					 XF.AddResult(e.toString(), "3");
					 XF.XmlscreenShot(false);
					 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
					 Caller.StringExecutionStatus.append("\n");
					 					
				}
			}



	// Method for storing Value
	public  void  storeIn() throws InterruptedException
	{
		try 
		{
			String key = inputData.split("=")[0].trim();	
			String storeValue = inputData.split("=")[1].trim();
			map.put(key, storeValue);

			reportSheetRowNo = reportSheetRowNo +1;

		}catch (Exception e)
		{
			e.getMessage();
			//System.out.println(e);
	
			
			
		}
	}


	public  boolean  isNumeric() throws InterruptedException

	{
	    boolean numeric = true;
		try 
		{
  
	            Double num = Double.parseDouble(inputData);
	        
	        

		}catch (Exception e)
		{
			numeric = false;
			e.getMessage();
			
			
			
		}
		return numeric;
	}

	// Method for report in customized report sheet
	public  void  report() throws InterruptedException
	{
		try 
		{
			Boolean dataFetchFlag = false;
			Boolean writingValues = false;
			String actualVal = "";
			String expectedVal = "";
			String result = "";
			String valueToReport = "";
			String noOfColHeaders[] = inputData.split(";");
			for(int eachReportVal = 0; eachReportVal<noOfColHeaders.length;eachReportVal++)
			{
				String reportColHeader = noOfColHeaders[eachReportVal].split("=")[0].trim();
				String valueToUpdate = noOfColHeaders[eachReportVal].split("=")[1].trim();

				if(valueToUpdate.contains("'"))
				{
					valueToReport = valueToUpdate.replace("'", "");
				}else
				{
					dataFetchFlag = CF.getDataFromDataSheet(valueToUpdate,"DataSheet");
					if(dataFetchFlag == true)
					{
						valueToReport = dataSheetValue;
					}else
					{
						try
						{
							valueToReport = map.get(valueToUpdate).toString();
						}catch (Exception e)
						{
							valueToReport = "";
						}
					}
				}

				if(reportColHeader.toUpperCase().startsWith("ACT"))
				{
					actualVal = valueToReport;
				}

				if(reportColHeader.toUpperCase().startsWith("EXP"))
				{
					expectedVal = valueToReport;
				}


				if(writingValues == false)
				{
					
				}
			}

			if(actualVal.equalsIgnoreCase(expectedVal))
			{
				result = "Pass";
			}else
			{
				result = "Fail";
			}

	

		}catch (Exception e)
		{
			e.getMessage();
			System.out.println(e);

			reportSheetRowNo = reportSheetRowNo +1;
			takeSnapShot = "y";
		}
	}



	// Method for get getTableRowCount
	public  String  getTableRowCount() throws InterruptedException
	{
		try 
		{
			WebElement table = null;
			if(objectType.equalsIgnoreCase("css"))
			{
				// Grab the table
				table = driver.findElement(By.cssSelector(objectProperty));
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				table = driver.findElement(By.xpath(objectProperty));
			}else if(objectType.equalsIgnoreCase("id"))
			{
				table = driver.findElement(By.id(objectProperty));
			}else if(objectType.equalsIgnoreCase("name"))
			{
				table = driver.findElement(By.name(objectProperty));
			}else if(objectType.equalsIgnoreCase("link"))
			{
				table = driver.findElement(By.linkText(objectProperty));
			}

			// Now get all the TR elements from the table
			List<WebElement> allRows = table.findElements(By.tagName("tr"));
			int rowCount = allRows.size() - 1;
			String rowCountStr = "" + rowCount;
			map.put(inputData, rowCountStr);

	

			return rowCountStr;

		}catch (Exception e)
		{
			e.getMessage();
			System.out.println(e);
			
		}
		return "";
	}



	// Method for get Table  Column Header names
	public  String  getTableColumnHeader() throws InterruptedException
	{
		try 
		{
			WebElement table = null;
			if(objectType.equalsIgnoreCase("css"))
			{
				// Grab the table
				table = driver.findElement(By.cssSelector(objectProperty));
			}else if(objectType.equalsIgnoreCase("xpath"))
			{
				table = driver.findElement(By.xpath(objectProperty));
			}else if(objectType.equalsIgnoreCase("id"))
			{
				table = driver.findElement(By.id(objectProperty));
			}else if(objectType.equalsIgnoreCase("name"))
			{
				table = driver.findElement(By.name(objectProperty));
			}else if(objectType.equalsIgnoreCase("link"))
			{
				table = driver.findElement(By.linkText(objectProperty));
			}

			String tblColumnHeaders = "";
			// Now get all the TR elements from the table
			List<WebElement> allRows = table.findElements(By.tagName("tr"));
			int rowCount = allRows.size();
			// And iterate over them, getting the cells
			for (WebElement row : allRows) 
			{
				List<WebElement> cells = row.findElements(By.tagName("th"));
				int cellsCount = cells.size();
				for (WebElement cell : cells)
				{
					String cellData = cell.getText();
					if(!cellData.equalsIgnoreCase(""))
					{
						if(tblColumnHeaders.equalsIgnoreCase(""))
						{
							tblColumnHeaders = cellData;
						}else
						{
							tblColumnHeaders = tblColumnHeaders + "; " + cellData;
						}
					}
				}
				if(cellsCount>0)
				{
					break;
				}
			}

			map.put(inputData, tblColumnHeaders);

		

			return tblColumnHeaders;

		}catch (Exception e)
		{
			e.getMessage();
			System.out.println(e);
	
		}
		return "";
	}




	// Method for find a cell data in a column of table then fetch the cell data of a diff column of that perticular row.
		public static  void ExcelPOSDashboardwrite(String FileName,String SheetName) {

        FileInputStream fis=null;
        try {
               fis= new FileInputStream(new File(FileName));
               XSSFWorkbook WorkBook= new XSSFWorkbook(fis);
            
                     XSSFSheet sheet= WorkBook.getSheet(SheetName);
                     XSSFCellStyle style1 = WorkBook.createCellStyle();
            	        style1.setBorderBottom(BorderStyle.THIN);
            	        style1.setBorderTop(BorderStyle.THIN);
            	        style1.setBorderRight(BorderStyle.THIN);
            	        style1.setBorderLeft(BorderStyle.THIN);
               
               int RowIndex = sheet.getPhysicalNumberOfRows();
               XSSFRow r=sheet.createRow(RowIndex);
               XSSFCell Serial = r.createCell(0);
               //String serialindex=Integer.toString(RowIndex);
               int serialno=RowIndex-2;
               Serial.setCellValue(serialno);
               Serial.setCellStyle(style1);
               XSSFCell startTime = r.createCell(1);
           	SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss" );
       	    String strstartTime = dateFormat.format(Caller.LOBStartTime); 
               startTime.setCellValue(strstartTime);
               startTime.setCellStyle(style1);

               XSSFCell endTime = r.createCell(2);
          	    String strendtime = dateFormat.format(Caller.LOBEndTime);
               endTime.setCellValue(strendtime);
               endTime.setCellStyle(style1);
               XSSFCell Duration = r.createCell(3);
               Duration.setCellValue(Caller.TotalTimeTaken);
               Duration.setCellStyle(style1);
               XSSFCell Lobname = r.createCell(4);
               Lobname.setCellValue(Caller.LOB);
               Lobname.setCellStyle(style1);
               
               XSSFCell testcasecount = r.createCell(5);
               testcasecount.setCellValue(Caller.LOBStatus);
               testcasecount.setCellStyle(style1);
               
               XSSFCell Testcasecount = r.createCell(6);
               Testcasecount.setCellValue(Caller.TotalTestCaseCount);
               Testcasecount.setCellStyle(style1);
               
                 
                            XSSFCell TotalExecutedTC = r.createCell(7);
                    
                            TotalExecutedTC.setCellValue(Caller.TotalExecuted);
                            TotalExecutedTC.setCellStyle(style1);
                                                     
                            XSSFCell TCPassed = r.createCell(8);
                            TCPassed.setCellValue(Caller.TotalPassed);
                            TCPassed.setCellStyle(style1);
//                  
                            XSSFCell TCFailed = r.createCell(9);
                            TCFailed.setCellValue(Caller.TotalFailed);
                            TCFailed.setCellStyle(style1);
                          
                            XSSFCell TCNotExe = r.createCell(10);
                            TCNotExe.setCellValue(Caller.TotalNotExecuted);
                            TCNotExe.setCellStyle(style1);
                            
                            XSSFCell TCPassPer = r.createCell(11);
                            TCPassPer.setCellValue(Caller.ExecutionPercentage);
                            TCPassPer.setCellStyle(style1);
                            
                            FileOutputStream outFile= new FileOutputStream(FileName);
                            WorkBook.write(outFile);
                     
                     
                            outFile.close();
                   
	}catch(IOException e) {
			e.printStackTrace();
			Caller.ExecutionStatus="FAIL";
			Caller.stringBuilder.append(e);
			Caller.stringBuilder.append("\n");
			
		}
	}
	public static void DateCalculation() throws Exception{
		
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss" );

		Calendar calendar = Calendar.getInstance();
	    String strstartTime = dateFormat.format(Caller.LOBStartTime); 
	    Caller.LOBEndTime = calendar.getTime();
	
		String strendTime = dateFormat.format(Caller.LOBEndTime);  
	    Date sttime=dateFormat.parse(strstartTime);
	    Date Endtime=dateFormat.parse(strendTime);
	    long diff =Endtime.getTime() - sttime.getTime();	
	    
	     Caller. TotalTimeTaken = diff / (60 * 1000) % 60;
	   
	   
	    
	}
	public static void TodayDate() throws Exception{
		
		SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");

		Calendar calendar = Calendar.getInstance();

	    Caller.CurrentTime = calendar.getTime();
	
		String strendTime = dateFormat.format(Caller.CurrentTime);  
	   

	    
	   
	   
		map.put(inputData, strendTime);
	}


	
	public static  void CreatePOSDashboardFile()  
	{
		
		
    	FileInputStream fis=null;
		try
		{

	        	 fis= new FileInputStream(new File(BSDFinalReportPath));
	             XSSFWorkbook  workbook = new XSSFWorkbook(fis);            
	            POSSheetName = Caller.BSDNo+"_Execution Dashboard";
	            XSSFSheet spreadsheet = workbook.createSheet(POSSheetName);
	            XSSFFont font = workbook.createFont();
	            font.setFontHeightInPoints((short) 30);
	            font.setFontName("Calibri");
	            font.setBold(true);
	            XSSFFont font1 = workbook.createFont();
	       
	            font1.setBold(true);
	        //XSSFCellStyle style=workbook.createCellStyle();
	        
	        XSSFCellStyle style1 = workbook.createCellStyle();
	        style1.setAlignment(XSSFCellStyle.ALIGN_LEFT);
	        style1.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        style1.setFont(font1);			
	        XSSFCellStyle style2 = workbook.createCellStyle();
	        style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
	        style2.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        style2.setFont(font1);
	        style2.setBorderBottom(BorderStyle.THIN);
	        style2.setBorderTop(BorderStyle.THIN);
	        style2.setBorderRight(BorderStyle.THIN);
	        style2.setBorderLeft(BorderStyle.THIN);
	        XSSFCellStyle style6 = workbook.createCellStyle();
	        style6.setAlignment(XSSFCellStyle.ALIGN_CENTER);
	        style6.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
	        style6.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        style6.setFont(font);
	        style6.setBorderBottom(BorderStyle.THIN);
	        style6.setBorderTop(BorderStyle.THIN);
	        style6.setBorderRight(BorderStyle.THIN);
	        style6.setBorderLeft(BorderStyle.THIN);
	        XSSFCellStyle style4 = workbook.createCellStyle();
	        style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
	        style4.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
	        style4.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        style4.setFont(font1);
	        style4.setBorderBottom(BorderStyle.THIN);
	        style4.setBorderTop(BorderStyle.THIN);
	        style4.setBorderRight(BorderStyle.THIN);
	        style4.setBorderLeft(BorderStyle.THIN);
	        XSSFRow row = spreadsheet.createRow(0);
	        XSSFCell  cell = (XSSFCell) row.createCell(0);
	        
	        spreadsheet.addMergedRegion(new CellRangeAddress(0,0,0,11));
	        cell.setCellValue("Execution Dashboard");
	      
	        //cell.setCellStyle(style6);
	        
	
	    
	        cell.setCellStyle(style6);

	        row = spreadsheet.createRow((short) 1);
	        cell = (XSSFCell) row.createCell(0);
	        spreadsheet.addMergedRegion(new CellRangeAddress(1,2,0,0));
	        cell.setCellValue("SL.NO");
	        cell.setCellStyle(style1);
	       
	        cell = (XSSFCell) row.createCell(1);
	        spreadsheet.addMergedRegion(new CellRangeAddress(1,2,1,1));
	        cell.setCellValue("Execution Start");
	        cell.setCellStyle(style1);

	        
	        cell = (XSSFCell) row.createCell(2);
	        spreadsheet.addMergedRegion(new CellRangeAddress(1,2,2,2));
	        cell.setCellValue("Execution End");
	        cell.setCellStyle(style1);
	        cell.setCellStyle(style1);
	        
	        cell = (XSSFCell) row.createCell(3);
	        spreadsheet.addMergedRegion(new CellRangeAddress(1,2,3,3));
	        cell.setCellValue("Duration");
	        cell.setCellStyle(style1);
	        cell.setCellStyle(style1);
	        

	        cell = (XSSFCell) row.createCell(4);
	        spreadsheet.addMergedRegion(new CellRangeAddress(1,2,4,4));

	        cell.setCellValue("LOB");
	        cell.setCellStyle(style1);
	        cell.setCellStyle(style1);
	        
	        cell = (XSSFCell) row.createCell(5);
	        spreadsheet.addMergedRegion(new CellRangeAddress(1,2,5,5));
	        cell.setCellValue("Status");
	        cell.setCellStyle(style1);
	        cell.setCellStyle(style1);
	       
	         cell = (XSSFCell) row.createCell(6);
	         cell.setCellValue("Test Case Execution Summary");
	         cell.setCellStyle(style2);
	         cell.setCellStyle(style2);

	        
	        spreadsheet.addMergedRegion(new CellRangeAddress(1,1,6,11));
	        row = spreadsheet.createRow((short) 2);
	        cell = (XSSFCell) row.createCell(6);
	        cell.setCellValue("Total cases");
	        cell.setCellStyle(style4);
	        cell = (XSSFCell) row.createCell(7);
	        cell.setCellValue("Executed");
	        cell.setCellStyle(style4);
	        
	       
	        cell = (XSSFCell) row.createCell(8);
	        cell.setCellValue("Passed");
	        cell.setCellStyle(style4);
	        cell = (XSSFCell) row.createCell(9);
	        cell.setCellValue("Failed");
	        cell.setCellStyle(style4);
	        cell = (XSSFCell) row.createCell(10);
	        cell.setCellValue("Not Executed");
	        cell.setCellStyle(style4);

	        cell = (XSSFCell) row.createCell(11);
	        cell.setCellValue("% Passed");
	        cell.setCellStyle(style4);
	        
	        setBordersToMergedCells(workbook,spreadsheet);
	        FileOutputStream fos = new FileOutputStream(BSDFinalReportPath);

	        spreadsheet.setColumnWidth(0, 2500);
	        spreadsheet.setColumnWidth(1, 2500);
	        spreadsheet.setColumnWidth(2, 2500);
	        spreadsheet.setColumnWidth(3, 2500);
	        spreadsheet.setColumnWidth(4, 2500);
	        spreadsheet.setColumnWidth(5, 2500);
	        spreadsheet.setColumnWidth(6, 2500);
	        spreadsheet.setColumnWidth(7, 2500);
	        spreadsheet.setColumnWidth(8, 2500);
	        spreadsheet.setColumnWidth(9, 2500);
	        spreadsheet.setColumnWidth(10, 2500);
	        spreadsheet.setColumnWidth(11, 2500);    
	        workbook.write(fos);
	        fos.close();
	        }       //System.out.println("typesofcells.xlsx written successfully");
		
	  
	
	catch(Exception e) {
		e.printStackTrace();
	}
	
	}
	

	
    public static  void ExcelResultwrite(String FileName,String SheetName) {

        FileInputStream fis=null;
        try {
        	
        	
               fis= new FileInputStream(new File(FileName));
               XSSFWorkbook WorkBook= new XSSFWorkbook(fis);
               XSSFCellStyle style1 = WorkBook.createCellStyle();
               style1.setBorderBottom(BorderStyle.THIN);
      	        style1.setBorderTop(BorderStyle.THIN);
      	        style1.setBorderRight(BorderStyle.THIN);
      	        style1.setBorderLeft(BorderStyle.THIN);
            
                     XSSFSheet sheet= WorkBook.getSheet(SheetName);
               
               int RowIndex = sheet.getPhysicalNumberOfRows();
               XSSFRow r=sheet.createRow(RowIndex);
               XSSFCell Serial = r.createCell(0);
               String serialindex=Integer.toString(RowIndex);
               Serial.setCellValue(serialindex);
               Serial.setCellStyle(style1);
               XSSFCell Lobnameindex = r.createCell(1);
               Lobnameindex.setCellValue(Caller.LOB);
               Lobnameindex.setCellStyle(style1);
               
               XSSFCell testcaseindex = r.createCell(2);
               testcaseindex.setCellValue(webActionKeywordsTest.testCaseID);
               testcaseindex.setCellStyle(style1);
               
               XSSFCell Executionstatus = r.createCell(3);
               String Execution = Caller.StringExecutionStatus.toString();
               
               if(Execution.contains("FAIL")) {
                            Executionstatus.setCellValue("FAILED");
                            Executionstatus.setCellStyle(style1);
                            Caller.TotalFailed = Caller.TotalFailed+1;
               }
               else {
                     Executionstatus.setCellValue("PASSED");
                     Executionstatus.setCellStyle(style1);
                     Caller.TotalPassed = Caller.TotalPassed+1;
               }
                                                        
                            
                            XSSFCell ActualActionResult = r.createCell(4);
                            String Validation=Caller.stringBuilder.toString();
                            ActualActionResult.setCellValue(Validation);
                            ActualActionResult.setCellStyle(style1);
                            XSSFCell BrResult = r.createCell(5);
                            String BR=Caller.BRRefernce.toString();
                            BrResult.setCellValue("These" +BR+"are Validated in this test case");
                            BrResult.setCellStyle(style1);
                            XSSFCellStyle hlink_style = WorkBook.createCellStyle();
                            CreationHelper createHelper = WorkBook.getCreationHelper();
                            Font hlink_font = WorkBook.createFont();
                            hlink_font.setUnderline(Font.U_SINGLE);
                            hlink_font.setColor(HSSFColor.BLUE.index);
                            hlink_style.setFont(hlink_font);
                            hlink_style.setBorderBottom(BorderStyle.THIN);
                            hlink_style.setBorderTop(BorderStyle.THIN);
                            hlink_style.setBorderRight(BorderStyle.THIN);
                            hlink_style.setBorderLeft(BorderStyle.THIN);
                          Hyperlink link = (XSSFHyperlink)createHelper.createHyperlink(Hyperlink.LINK_FILE);                                                      
                            XSSFCell TCScreenshotLink = r.createCell(6);
                            String TCScreenshotpath= BSDFinalScreenPath;
                        if(TCScreenshotpath!=null) {
                            TCScreenshotLink.setCellValue("Click Here");
                  
                            TCScreenshotpath=TCScreenshotpath.replace("\\", "/");
                            File NewPath = new File(TCScreenshotpath);
                            link.setAddress(NewPath.toURI().toString());
                            
                            TCScreenshotLink.setHyperlink(link);
                            TCScreenshotLink.setCellStyle(hlink_style);
                         
//                            link.setAddress(TCScreenshotpath);                         
//                            TCScreenshotLink.setHyperlink(link);    
//                            TCScreenshotLink.setCellStyle(hlink_style);
                              }
else {
System.out.println("The screenshot is not available");	
}
                          
                            
                            FileOutputStream outFile= new FileOutputStream(FileName);
                            WorkBook.write(outFile);
                     
                     
                            outFile.close();
                   
	}catch(IOException e) {
			e.printStackTrace();
			Caller.ExecutionStatus="FAIL";
			Caller.stringBuilder.append(e);
			Caller.stringBuilder.append("\n");
			
		}
	}

    public static  void CreateExcelFile()
    {
    	try {
    	
    	SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");
    	Calendar calendar = Calendar.getInstance();
    	Caller.startDate = calendar.getTime();
        String strstartTime = dateFormat.format(Caller.startDate); 
        String Tempstring = strstartTime.replace(":","");
    	String TempDate[]=strstartTime.split(" ",2);
    	BSDScreenshotPath=Caller.BSDExecutionResultPath+"\\"+TempDate[0]+"\\"+Caller.BSDNo+"\\";
    	// Templatepath+ResultsFolder+BSDNAme+ResultExcel
        BSDFinalResultPath =BSDScreenshotPath+Caller.BSDNo+"_Result_"+Tempstring+".xlsx";
    	
    	//BSDFinalResultPath ="D:\\MMAProject\\MMA_Automation\\Results"+"\\"+TempDate[0]+"\\"+"BSD_013"+"\\"+"BSD_013"+"_Result_"+Tempstring+".xlsx";
        File f= new File(BSDFinalResultPath);
        if(!f.exists()&& (!f.isDirectory())) {
        	f.getParentFile().mkdirs();
        	f.createNewFile();
            FileOutputStream fos = new FileOutputStream(BSDFinalResultPath);
            XSSFWorkbook  workbook = new XSSFWorkbook();            
            XSSFSheet sheet = workbook.createSheet("Result Sheet");
          	 final Font font = sheet.getWorkbook ().createFont ();
     	    font.setFontName ("Arial");
     	    font.setBoldweight (Font.BOLDWEIGHT_BOLD );
     	    font.setColor ( HSSFColor.BLACK.index );
            XSSFCellStyle style = workbook.createCellStyle();
     	    style.setFont ( font );
     	    style.setFillForegroundColor ( HSSFColor.LIGHT_ORANGE.index);
     	    style.setFillPattern ( PatternFormatting.SOLID_FOREGROUND );
     	    style.setBorderBottom(BorderStyle.THIN);
  	        style.setBorderTop(BorderStyle.THIN);
  	        style.setBorderRight(BorderStyle.THIN);
  	        style.setBorderLeft(BorderStyle.THIN);
            XSSFRow row = sheet.createRow(0);   
            XSSFCell SerialNo = row.createCell(0);
            SerialNo.setCellValue("S.No");
            SerialNo.setCellStyle(style);
            XSSFCell LOBName = row.createCell(1);
            LOBName.setCellValue("Line of Business");
            LOBName.setCellStyle(style);
            XSSFCell TestCaseName = row.createCell(2);
            TestCaseName.setCellValue("Test Case Name");
            TestCaseName.setCellStyle(style);
            XSSFCell TCResult = row.createCell(3);
            TCResult.setCellValue("Result");
            TCResult.setCellStyle(style);
            XSSFCell TCReport = row.createCell(4);
            TCReport.setCellValue("Execution Report");
            TCReport.setCellStyle(style);
            XSSFCell TCBusinessRule = row.createCell(5);
            TCBusinessRule.setCellValue("Business Rule");
            TCBusinessRule.setCellStyle(style);
            XSSFCell TCScreenshot = row.createCell(6);
            TCScreenshot.setCellValue("ScreenShot Link");
            TCScreenshot.setCellStyle(style);
            sheet.setColumnWidth(0, 1500);
            sheet.setColumnWidth(1, 5000);
            sheet.setColumnWidth(2, 6000);
            sheet.setColumnWidth(3, 3500);
            sheet.setColumnWidth(4, 20000);
            sheet.setColumnWidth(5, 5000);
            sheet.setColumnWidth(6, 5000);
            
            
            workbook.write(fos);
       
            fos.close();
        }
    	}
        catch(Exception e){
        	e.printStackTrace();
        }
    	
    
    }
    public static void screenShot() throws IOException, InterruptedException
    { 
    	CreateFolder();
     	SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");
    	Calendar calendar = Calendar.getInstance();
    	Caller.screenshottime = calendar.getTime();
        String strstartTime = dateFormat.format(Caller.screenshottime); 
        String Tempstring = strstartTime.replace(":","");
    	String TempDate[]=Tempstring.split(" ",2);
        JavascriptExecutor jexec = (JavascriptExecutor)webActionKeywordsTest.driver;
        jexec.executeScript("window.scrollTo(0,0)"); // will scroll to (0,0) position 
        boolean isScrollBarPresent = (boolean)jexec.executeScript("return document.body.scrollHeight>document.body.clientHeight");
        long scrollHeight = (long)jexec.executeScript("return document.body.scrollHeight");

        long clientHeight = (long)jexec.executeScript("return document.body.clientHeight");
        webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
			if(webActionKeywordsTest.PageName.contains("/")) {
				webActionKeywordsTest.PageName.replace("/","_");
			}
        int fileIndex = 1;
        if(webActionKeywordsTest.driver instanceof InternetExplorerDriver){
//            if(isScrollBarPresent){
//                while(scrollHeight > 0){
//                	
//                File srcFile = ((TakesScreenshot)webActionKeywordsTest.driver).getScreenshotAs(OutputType.FILE);
//                    org.apache.commons.io.FileUtils.copyFile(srcFile, new File(BSDFinalScreenPath+webActionKeywordsTest.PageName+"_"+webActionKeywordsTest.ReportObjectName+fileIndex+TempDate[0]+".jpg"));
//               
//                    jexec.executeScript("window.scrollTo(0,"+clientHeight*fileIndex++ +")");
//                    scrollHeight = scrollHeight - clientHeight;
//                   
//        			 Caller.ExecutionStatus="PASS";	      
//        			 Caller.Reason=webActionKeywordsTest.action+"  action is Executed SuccessFully ";
//        			 Caller.stringBuilder.append(Caller.Reason);	
//        			 Caller.stringBuilder.append("\n");
//        			 
//        			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
//        			 Caller.StringExecutionStatus.append("\n");
//
//                   
//                                      
//                                           
//            }
//            }else{
                File srcFile = ((TakesScreenshot)webActionKeywordsTest.driver).getScreenshotAs(OutputType.FILE);
                org.apache.commons.io.FileUtils.copyFile(srcFile, new File(BSDFinalScreenPath+webActionKeywordsTest.PageName+"_"+webActionKeywordsTest.ReportObjectName+TempDate[0]+".jpg"));
   			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason=webActionKeywordsTest.action+"  action is Executed SuccessFully ";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");

//            }
        }else{
            File srcFile = ((TakesScreenshot)webActionKeywordsTest.driver).getScreenshotAs(OutputType.FILE);
            org.apache.commons.io.FileUtils.copyFile(srcFile, new File(BSDFinalScreenPath+webActionKeywordsTest.PageName+"_"+webActionKeywordsTest.ReportObjectName+TempDate[0]+".jpg"));
			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason=webActionKeywordsTest.action+"  action is Executed SuccessFully ";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");

        }
        

    }





// Method for find a cell data in a column of table then fetch the cell data of a diff column of that perticular row and compare with expected cell value.
public  String  getTableCellTxtCompare() throws InterruptedException
{
	return null;


}


public static void setBordersToMergedCells(XSSFWorkbook workBook, XSSFSheet sheet) {
    int numMerged = sheet.getNumMergedRegions();

for(int i= 0; i<numMerged;i++){
    CellRangeAddress mergedRegions = sheet.getMergedRegion(i);
    RegionUtil.setBorderTop(CellStyle.BORDER_THIN, mergedRegions, sheet, workBook);
    RegionUtil.setBorderLeft(CellStyle.BORDER_THIN, mergedRegions, sheet, workBook);
    RegionUtil.setBorderRight(CellStyle.BORDER_THIN, mergedRegions, sheet, workBook);
    RegionUtil.setBorderBottom(CellStyle.BORDER_THIN, mergedRegions, sheet, workBook);
}


}


public static int getNbOfMergedRegions(int row,String Col_Address) throws IOException

{
	int temprow=row+1;
	String RowAddress=Integer.toString(temprow);
	
	FileInputStream fis=null;
 	 fis= new FileInputStream(new File(Caller.UtilitySheetpath));

	 XSSFWorkbook workbook = new XSSFWorkbook(fis);
	//XSSFWorkbook WorkBook= new XSSFWorkbook();
	XSSFSheet NewSheet= workbook.getSheet("Launch Sheet");
	 int value = 0 ;
    for(int i = 0; i < NewSheet.getNumMergedRegions(); i++)
    {
        CellRangeAddress range = NewSheet.getMergedRegion(i);

   
        String RangeValue=range.toString();
         String[] temp=   RangeValue.split(" ",2);
       
         if(temp[1].contains(RowAddress)&& temp[1].contains(Col_Address))
         {
	     value= range.getNumberOfCells(); 
	     break;
	      //System.out.println(value);
          }
     }
    
    if (value == 0){
    	value = 1;
    }
    return value;
    
}

public static  void CreateSummaryReport()
{
	try {
	
	SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");
	Calendar calendar = Calendar.getInstance();
	Caller.startDate = calendar.getTime();
    String strstartTime = dateFormat.format(Caller.startDate); 
    String Tempstring = strstartTime.replace(":","");
	String TempDate[]=Tempstring.split(" ",2);
	BSDSummaryReportPath=Caller.BSDExecutionResultPath+"\\"+TempDate[0];
	
    File oldFile= new File(BSDSummaryReportPath);
	 if((!oldFile.isDirectory())) {
		 oldFile.mkdir();
	    }


	 XSSFWorkbook workbook = new XSSFWorkbook();
	 FileOutputStream out = new FileOutputStream(new File(oldFile+"\\"+"createBlankWorkBook.xlsx"));


	      workbook.write(out);
	      
    BSDFinalReportPath =oldFile+"\\"+"createBlankWorkBook.xlsx";
    
    CF.copyPasteFile(Caller.TestSummaryReportPath, BSDFinalReportPath);

    File TempFile=new File(BSDFinalReportPath);
    File NewFile=new File(oldFile+"\\"+"POS_Test Summary Report"+"_"+TempDate[0]+"_"+TempDate[1]+".xlsx");
    TempFile.renameTo(NewFile);

    BSDFinalReportPath=oldFile+"\\"+"POS_Test Summary Report"+"_"+TempDate[0]+"_"+TempDate[1]+".xlsx";
    out.close();
	}
    catch(Exception e){
    	e.printStackTrace();
    }
	

}
public static  void CreateSmokeTestingResultExcelFile()
{
	try {
	
	SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");
	Calendar calendar = Calendar.getInstance();
	Caller.startDate  = calendar.getTime();
    String strstartTime = dateFormat.format(Caller.startDate); 
    String Tempstring = strstartTime.replace(":","");
	String TempDate[]=strstartTime.split(" ",2);
	BSDScreenshotPath=Caller.BSDExecutionResultPath+"\\"+TempDate[0]+"\\Smoke Testing\\";
	// Templatepath+ResultsFolder+BSDNAme+ResultExcel
    BSDFinalResultPath =BSDScreenshotPath+"Smoke Testing Results_"+Tempstring+".xlsx";
	
	//BSDFinalResultPath ="D:\\MMAProject\\MMA_Automation\\Results"+"\\"+TempDate[0]+"\\"+"BSD_013"+"\\"+"BSD_013"+"_Result_"+Tempstring+".xlsx";
    File f= new File(BSDFinalResultPath);
    if(!f.exists()&& (!f.isDirectory())) {
    	f.getParentFile().mkdirs();
    	f.createNewFile();
        FileOutputStream fos = new FileOutputStream(BSDFinalResultPath);
        XSSFWorkbook  workbook = new XSSFWorkbook();            
        XSSFSheet sheet = workbook.createSheet("Result Sheet");
      	 final Font font = sheet.getWorkbook ().createFont ();
 	    font.setFontName ("Arial");
 	    font.setBoldweight (Font.BOLDWEIGHT_BOLD );
 	    font.setColor ( HSSFColor.BLACK.index );
        XSSFCellStyle style = workbook.createCellStyle();
 	    style.setFont ( font );
 	    style.setFillForegroundColor ( HSSFColor.LIGHT_ORANGE.index);
 	    style.setFillPattern ( PatternFormatting.SOLID_FOREGROUND );
 	    style.setBorderBottom(BorderStyle.THIN);
	        style.setBorderTop(BorderStyle.THIN);
	        style.setBorderRight(BorderStyle.THIN);
	        style.setBorderLeft(BorderStyle.THIN);
        XSSFRow row = sheet.createRow(0);   
        XSSFCell SerialNo = row.createCell(0);
        SerialNo.setCellValue("S.No");
        SerialNo.setCellStyle(style);
//        XSSFCell LOBName = row.createCell(1);
//        LOBName.setCellValue("Line of Business");
//        LOBName.setCellStyle(style);
        XSSFCell TestCaseName = row.createCell(1);
        TestCaseName.setCellValue("POS No");
        TestCaseName.setCellStyle(style);
        XSSFCell TCResult = row.createCell(2);
        TCResult.setCellValue("Status");
        TCResult.setCellStyle(style);
        XSSFCell TCReport = row.createCell(3);
        TCReport.setCellValue("Execution Report");
        TCReport.setCellStyle(style);
        XSSFCell TCScreenshot = row.createCell(4);
        TCScreenshot.setCellValue("ScreenShots");
        TCScreenshot.setCellStyle(style);
        sheet.setColumnWidth(0, 1500);
        sheet.setColumnWidth(1, 7000);
        sheet.setColumnWidth(2, 3500);
        sheet.setColumnWidth(3, 20000);
        sheet.setColumnWidth(4, 3500);
//        sheet.setColumnWidth(5, 5000);
        
        
        workbook.write(fos);
   
        fos.close();
    }
	}
    catch(Exception e){
    	e.printStackTrace();
    }
	

}





public static  void CreateSmokeDashboardFile()  
{
	
	
	FileInputStream fis=null;
	try
	{

        	 fis= new FileInputStream(new File(BSDFinalResultPath));
             XSSFWorkbook  workbook = new XSSFWorkbook(fis);            
            POSSheetName = "Smoke Testing Dashboard";
            XSSFSheet spreadsheet = workbook.createSheet(POSSheetName);
            XSSFFont font = workbook.createFont();
            font.setFontHeightInPoints((short) 16);
            font.setFontName("Calibri");
            font.setBold(true);
            
            XSSFFont font1 = workbook.createFont();
       
            font1.setBold(true);
        //XSSFCellStyle style=workbook.createCellStyle();
        
        XSSFCellStyle style1 = workbook.createCellStyle();
        style1.setAlignment(XSSFCellStyle.ALIGN_LEFT);
        style1.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style1.setFont(font1);			
        XSSFCellStyle style2 = workbook.createCellStyle();
        style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        style2.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style2.setFont(font1);
        style2.setBorderBottom(BorderStyle.THIN);
        style2.setBorderTop(BorderStyle.THIN);
        style2.setBorderRight(BorderStyle.THIN);
        style2.setBorderLeft(BorderStyle.THIN);
        XSSFCellStyle style6 = workbook.createCellStyle();
        style6.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        style6.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        style6.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style6.setFont(font);
        style6.setBorderBottom(BorderStyle.THIN);
        style6.setBorderTop(BorderStyle.THIN);
        style6.setBorderRight(BorderStyle.THIN);
        style6.setBorderLeft(BorderStyle.THIN);
        XSSFCellStyle style4 = workbook.createCellStyle();
        style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        style4.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style4.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style4.setFont(font1);
        style4.setBorderBottom(BorderStyle.THIN);
        style4.setBorderTop(BorderStyle.THIN);
        style4.setBorderRight(BorderStyle.THIN);
        style4.setBorderLeft(BorderStyle.THIN);
        XSSFRow row = spreadsheet.createRow(0);
        XSSFCell  cell = (XSSFCell) row.createCell(0);
        
        spreadsheet.addMergedRegion(new CellRangeAddress(0,0,0,4));
        

   	    
   	SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
 	Calendar calendar = Calendar.getInstance();
 	Caller.startDate = calendar.getTime();
    String strstartTime = dateFormat.format(Caller.startDate);   	    
    cell.setCellValue("Smoke Testing - Execution Dashboard as on - "+strstartTime);
      
        //cell.setCellStyle(style6);
        

    
        cell.setCellStyle(style6);

//        row = spreadsheet.createRow((short) 1);
//        cell = (XSSFCell) row.createCell(0);
//        spreadsheet.addMergedRegion(new CellRangeAddress(1,2,0,0));
//        cell.setCellValue("Execution Start");
//        cell.setCellStyle(style1);
//       
//        cell = (XSSFCell) row.createCell(1);
//        spreadsheet.addMergedRegion(new CellRangeAddress(1,2,1,1));
//        cell.setCellValue("Execution End");
//        cell.setCellStyle(style1);
//
//        
//        cell = (XSSFCell) row.createCell(2);
//        spreadsheet.addMergedRegion(new CellRangeAddress(1,2,2,2));
//        cell.setCellValue("Duration");
//        cell.setCellStyle(style1);
//        cell.setCellStyle(style1);
//        
//        cell = (XSSFCell) row.createCell(3);
//        spreadsheet.addMergedRegion(new CellRangeAddress(1,2,3,3));
//        cell.setCellValue("Status");
//        cell.setCellStyle(style1);
//        cell.setCellStyle(style1);
        

//        cell = (XSSFCell) row.createCell(4);
//        spreadsheet.addMergedRegion(new CellRangeAddress(1,2,4,4));
//
//        cell.setCellValue("LOB");
//        cell.setCellStyle(style1);
//        cell.setCellStyle(style1);
//        
//        cell = (XSSFCell) row.createCell(5);
//        spreadsheet.addMergedRegion(new CellRangeAddress(1,2,5,5));
//        cell.setCellValue("Status");
//        cell.setCellStyle(style1);
//        cell.setCellStyle(style1);
       
//         cell = (XSSFCell) row.createCell(0);
//         cell.setCellValue("Execution Status Summary");
//         cell.setCellStyle(style2);
//         cell.setCellStyle(style2);

        
//        spreadsheet.addMergedRegion(new CellRangeAddress(1,1,4,6));
        row = spreadsheet.createRow((short) 1);
        cell = (XSSFCell) row.createCell(0);
        cell.setCellValue("Total POS");
        cell.setCellStyle(style4);
        cell = (XSSFCell) row.createCell(1);
        cell.setCellValue("Executed");
        cell.setCellStyle(style4);
        cell = (XSSFCell) row.createCell(2);
        cell.setCellValue("Passed");
        cell.setCellStyle(style4);
        cell = (XSSFCell) row.createCell(3);
        cell.setCellValue("Failed");
        cell.setCellStyle(style4);
        cell = (XSSFCell) row.createCell(4);
        cell.setCellValue("Not Executed");
        cell.setCellStyle(style4);
        
     

//        cell = (XSSFCell) row.createCell(5);
//        cell.setCellValue("% Passed");
//        cell.setCellStyle(style4);
        
      
        
        setBordersToMergedCells(workbook,spreadsheet);
        FileOutputStream fos = new FileOutputStream(BSDFinalResultPath);

        spreadsheet.setColumnWidth(0, 4000);
        spreadsheet.setColumnWidth(1, 4000);
        spreadsheet.setColumnWidth(2, 4000);
        spreadsheet.setColumnWidth(3, 4000);
        spreadsheet.setColumnWidth(4, 4000);
        spreadsheet.setColumnWidth(5, 4000);
//        spreadsheet.setColumnWidth(6, 2500);
//        spreadsheet.setColumnWidth(7, 2500);
//        spreadsheet.setColumnWidth(8, 2500);
//        spreadsheet.setColumnWidth(9, 2500);
//        spreadsheet.setColumnWidth(10, 2500);
//        spreadsheet.setColumnWidth(11, 2500);    
        workbook.write(fos);
        fos.close();
        }       //System.out.println("typesofcells.xlsx written successfully");
	
  

catch(Exception e) {
	e.printStackTrace();
}

}





public static  void ExcelSmokeDashboardwrite(String FileName,String SheetName) {

        FileInputStream fis=null;
        try {
               fis= new FileInputStream(new File(FileName));
               XSSFWorkbook WorkBook= new XSSFWorkbook(fis);
            
                     XSSFSheet sheet= WorkBook.getSheet(SheetName);
                     XSSFCellStyle style1 = WorkBook.createCellStyle();
                     XSSFFont font = WorkBook.createFont();
        	            font.setFontHeightInPoints((short) 16);
        	            font.setFontName("Calibri");
        	            font.setBold(true);
//        	            font.setColor(XSSFFont.COLOR_RED);
            	        style1.setBorderBottom(BorderStyle.THIN);
            	        style1.setBorderTop(BorderStyle.THIN);
            	        style1.setBorderRight(BorderStyle.THIN);
            	        style1.setBorderLeft(BorderStyle.THIN);
            	        style1.setAlignment(XSSFCellStyle.ALIGN_CENTER);
            	        style1.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            	        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            	        style1.setFont(font);	
            	       
        	            
        	            
               int RowIndex = sheet.getPhysicalNumberOfRows();
               XSSFRow r=sheet.createRow(RowIndex);
//               XSSFCell Serial = r.createCell(0);
//               //String serialindex=Integer.toString(RowIndex);
//               int serialno=RowIndex-2;
//               Serial.setCellValue(serialno);
//               Serial.setCellStyle(style1);
//               XSSFCell startTime = r.createCell(0);
//           	SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
//       	    String strstartTime = dateFormat.format(Caller.LOBStartTime); 
//               startTime.setCellValue(strstartTime);
//               startTime.setCellStyle(style1);
//
//               XSSFCell endTime = r.createCell(1);
//          	    String strendtime = dateFormat.format(Smoke_Testing_Suite.LOBEndTime);
//               endTime.setCellValue(strendtime);
//               endTime.setCellStyle(style1);
//               XSSFCell Duration = r.createCell(2);
//               Duration.setCellValue(Smoke_Testing_Suite.TotalTimeTaken);
//               Duration.setCellStyle(style1);
//               XSSFCell Lobname = r.createCell(4);
//               Lobname.setCellValue(Smoke_Testing_Suite.LOB);
//               Lobname.setCellStyle(style1);
               
//               XSSFCell testcasecount = r.createCell(3);
//               testcasecount.setCellValue(Smoke_Testing_Suite.LOBStatus);
//               testcasecount.setCellStyle(style1);
               
               XSSFCell Testcasecount = r.createCell(0);
               Testcasecount.setCellValue(Caller.TotalTestCaseCount);
               Testcasecount.setCellStyle(style1);
               
                 
                            XSSFCell TotalExecutedTC = r.createCell(1);
                    
                            TotalExecutedTC.setCellValue(Caller.TotalExecuted);
                            TotalExecutedTC.setCellStyle(style1);
                                                     
                            XSSFCell TCPassed = r.createCell(2);
                            TCPassed.setCellValue(Caller.TotalPassed);
                            TCPassed.setCellStyle(style1);
//                  
                            XSSFCell TCFailed = r.createCell(3);
                            TCFailed.setCellValue(Caller.TotalFailed);
                            TCFailed.setCellStyle(style1);
                          
                            XSSFCell TCNotExe = r.createCell(4);
                            TCNotExe.setCellValue(Caller.TotalNotExecuted);
                            TCNotExe.setCellStyle(style1);
                            
//                            XSSFCell TCPassPer = r.createCell(5);
//                            TCPassPer.setCellValue(Caller.ExecutionPercentage);
//                            TCPassPer.setCellStyle(style1);
                            
                            
                            r=sheet.createRow(RowIndex+1);
                            
                            XSSFCell Percentage = r.createCell(0);
                            Percentage.setCellValue("%");
                            Percentage.setCellStyle(style1);
                            
         	               
         	                 
         	                            XSSFCell ExecutedTCPer = r.createCell(1);
         	                            
//         	                          int ExecutedPer = Caller.TotalExecuted/Caller.TotalTestCaseCount*100;
//         	                    
//         	                           ExecutedTCPer.setCellValue(ExecutedPer);
         	                          ExecutedTCPer.setCellStyle(style1);
         	                                                     
         	                            XSSFCell TCPassedPer = r.createCell(2);
         	                           TCPassedPer.setCellValue(Caller.ExecutionPercentage+" %");
         	                          TCPassedPer.setCellStyle(style1);
//         	                  
         	                            XSSFCell TCFailedPer = r.createCell(3);
         	                            int FailedPer = (100 - Caller.ExecutionPercentage);
         	                           TCFailedPer.setCellValue(FailedPer+" %");
         	                          TCFailedPer.setCellStyle(style1);
         	                          
         	                            XSSFCell TCNotExePer = r.createCell(4);
//         	                           int NotExePer =  (100 - ExecutedPer);
//         	                           TCNotExePer.setCellValue(NotExePer);
         	                          TCNotExePer.setCellStyle(style1);
         	                            
//         	                            XSSFCell TCPassPer = r.createCell(5);
//      	                            TCPassPer.setCellValue(Caller.ExecutionPercentage);
//         	                            TCPassPer.setCellStyle(style1);
                            
                            
                            
                            FileOutputStream outFile= new FileOutputStream(FileName);
                            WorkBook.write(outFile);
                     
                     
                            outFile.close();
                   
	}catch(IOException e) {
			e.printStackTrace();
			Caller.ExecutionStatus="FAIL";
			Caller.stringBuilder.append(e);
			Caller.stringBuilder.append("\n");
			
		}
	}

public static  void ExcelResultwriteSmoke(String FileName,String SheetName) {

    FileInputStream fis=null;
    try {
         
         
           fis= new FileInputStream(new File(FileName));
           XSSFWorkbook WorkBook= new XSSFWorkbook(fis);
           XSSFCellStyle style1 = WorkBook.createCellStyle();
           style1.setBorderBottom(BorderStyle.THIN);
          style1.setBorderTop(BorderStyle.THIN);
          style1.setBorderRight(BorderStyle.THIN);
          style1.setBorderLeft(BorderStyle.THIN);
        
                 XSSFSheet sheet= WorkBook.getSheet(SheetName);
           
           int RowIndex = sheet.getPhysicalNumberOfRows();
           XSSFRow r=sheet.createRow(RowIndex);
           XSSFCell Serial = r.createCell(0);
           String serialindex=Integer.toString(RowIndex);
           Serial.setCellValue(serialindex);
           Serial.setCellStyle(style1);
//           XSSFCell Lobnameindex = r.createCell(1);
//           Lobnameindex.setCellValue(Caller.LOB);
//           Lobnameindex.setCellStyle(style1);
           
           XSSFCell testcaseindex = r.createCell(1);
           testcaseindex.setCellValue(webActionKeywordsTest.testCaseID);
           testcaseindex.setCellStyle(style1);
           
           XSSFCell Executionstatus = r.createCell(2);
           String Execution = Caller.StringExecutionStatus.toString();
           
           if(Execution.contains("FAIL")) {
                        Executionstatus.setCellValue("FAILED");
                        Executionstatus.setCellStyle(style1);
                        Caller.TotalFailed = Caller.TotalFailed+1;
           }
           else {
                 Executionstatus.setCellValue("PASSED");
                 Executionstatus.setCellStyle(style1);
                 Caller.TotalPassed = Caller.TotalPassed+1;
           }
                                                    
                        
                        XSSFCell ActualActionResult = r.createCell(3);
                        String Validation=Caller.stringBuilder.toString();
                        ActualActionResult.setCellValue(Validation);
                        ActualActionResult.setCellStyle(style1);
                        XSSFCellStyle hlink_style = WorkBook.createCellStyle();
                        CreationHelper createHelper = WorkBook.getCreationHelper();
                        Font hlink_font = WorkBook.createFont();
                        hlink_font.setUnderline(Font.U_SINGLE);
                        hlink_font.setColor(HSSFColor.BLUE.index);
                        hlink_style.setFont(hlink_font);
                        hlink_style.setBorderBottom(BorderStyle.THIN);
                        hlink_style.setBorderTop(BorderStyle.THIN);
                        hlink_style.setBorderRight(BorderStyle.THIN);
                        hlink_style.setBorderLeft(BorderStyle.THIN);
                      Hyperlink link = (XSSFHyperlink)createHelper.createHyperlink(Hyperlink.LINK_FILE);                                                      
                        XSSFCell TCScreenshotLink = r.createCell(4);
                        String TCScreenshotpath= BSDFinalScreenPath;

                        if(TCScreenshotpath!=null) {
                            TCScreenshotLink.setCellValue("Click Here");
                  
                            TCScreenshotpath=TCScreenshotpath.replace("\\", "/");
                            File NewPath = new File(TCScreenshotpath);
                            link.setAddress(NewPath.toURI().toString());
                            
                            TCScreenshotLink.setHyperlink(link);
                            TCScreenshotLink.setCellStyle(hlink_style);
                         
//                            link.setAddress(TCScreenshotpath);                         
//                            TCScreenshotLink.setHyperlink(link);    
//                            TCScreenshotLink.setCellStyle(hlink_style);
                              }
else {
System.out.println("The screenshot is not available");	
}                      
                      
                        
                        FileOutputStream outFile= new FileOutputStream(FileName);
                        WorkBook.write(outFile);
                 
                 
                        outFile.close();
               
   }catch(IOException e) {
                 e.printStackTrace();
                 Caller.ExecutionStatus="FAIL";
                 Caller.stringBuilder.append(e);
                 Caller.stringBuilder.append("\n");
                 
          }
   }



public static  void CreateFolder()
{
	try {

	// Templatepath+ResultsFolder+BSDNAme+ResultExcel
    BSDFinalScreenPath=BSDScreenshotPath+"ScreenShots\\"+Caller.LOB+"\\"+webActionKeywordsTest.testCaseID+"\\";
	//BSDFinalResultPath ="D:\\MMAProject\\MMA_Automation\\Results"+"\\"+TempDate[0]+"\\"+"BSD_013"+"\\"+"BSD_013"+"_Result_"+Tempstring+".xlsx";
    File f= new File(BSDFinalScreenPath);
    if((!f.isDirectory())) {
    	f.getParentFile().mkdirs();
    }

}catch(Exception e) {
	
}
}
}


