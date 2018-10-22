package businessKeywords;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

import allocator.Caller;

import allocator.CommonFunctions;
import allocator.Xml_Reporting_Functions;
import allocator.Caller;

public class BusinessFunctions {
	public static CommonFunctions CF = new CommonFunctions();
	public static Xml_Reporting_Functions XF = new Xml_Reporting_Functions();
	public static webActionKeywordsTest WF= new webActionKeywordsTest();
	public static Caller CA=new Caller();
	public static String testCaseID = "";
	public static String testFlowSheetName = "";
	public static WebDriver driver =  null;
	public static String objectName= "";
	public static String objectNameGiven= "";
	public static String action= "";
	public static String actionRowNo= "";
	public static String inputData= "";
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

	public static Map<String, String> map = new HashMap<String, String>();
	public static int AutomationStepCount;

	public  void  LoginEbao() throws InterruptedException, IOException
	{

		try 
		{
			 //USername
		CF.getMappedObjPropFromObjSheetByDB("UserName","RowNum");
		CF.getObjectProperty_Type(webActionKeywordsTest.objectName);
		webActionKeywordsTest.inputData=CA.Username;
		webActionKeywordsTest.ReportObjectName="UserName";
		WF.setText();
		//Password
		CF.getMappedObjPropFromObjSheetByDB("Password","RowNum");
		CF.getObjectProperty_Type(webActionKeywordsTest.objectName);
		webActionKeywordsTest.inputData=CA.Password;
		webActionKeywordsTest.ReportObjectName="Password";
		WF.setText();
		//Click
		CF.getMappedObjPropFromObjSheetByDB("Login","RowNum");
		CF.getObjectProperty_Type(webActionKeywordsTest.objectName);
		webActionKeywordsTest.ReportObjectName="Login";
		WF.click();
		webActionKeywordsTest.inputData="2";
		WF.sleep();
		//Verify
		CF.getMappedObjPropFromObjSheetByDB("POS","RowNum");
		CF.getObjectProperty_Type(webActionKeywordsTest.objectName);
		webActionKeywordsTest.ReportObjectName="POS";
	
		 Caller.ExecutionStatus="PASS";	      
		 Caller.Reason=webActionKeywordsTest.action+"has been Executed Successfully"; 
		 Caller.stringBuilder.append(Caller.Reason);	
		 Caller.stringBuilder.append("\n");
		 XF.AddResult(Caller.Reason, "1");
		 XF.XmlscreenShot(true);
		 
		 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
		 Caller.StringExecutionStatus.append("\n");

		} catch (Exception e)
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
	public  void  PreCapture() throws InterruptedException, IOException
	{

		try 
		{
			 //USername
		CF.getMappedObjPropFromObjSheetByDB("POS","RowNum");
		CF.getObjectProperty_Type(webActionKeywordsTest.objectName);
		
		webActionKeywordsTest.ReportObjectName="POS";
		WF.click();
		//Password
		CF.getMappedObjPropFromObjSheetByDB("PreCapture","RowNum");
		CF.getObjectProperty_Type(webActionKeywordsTest.objectName);

		webActionKeywordsTest.ReportObjectName="PreCapture";
		WF.click();
		
		webActionKeywordsTest.inputData="3";
		WF.sleep();
		
		//Enter Policy Number
		CF.getMappedObjPropFromObjSheetByDB("Policynumber","RowNum");
		CF.getObjectProperty_Type(webActionKeywordsTest.objectName);
		webActionKeywordsTest.ReportObjectName="Policynumber";
		webActionKeywordsTest.inputData="PolicyNo";
	    CF.getDataFromDataSheet(webActionKeywordsTest.inputData, "DataSheet");
	    webActionKeywordsTest.inputData = webActionKeywordsTest.dataSheetValue;
	    WF.setText();
	    //Select capture code
	    CF.getMappedObjPropFromObjSheetByDB("CaptureCode","RowNum");
		CF.getObjectProperty_Type(webActionKeywordsTest.objectName);
		webActionKeywordsTest.ReportObjectName="CaptureCode";
		webActionKeywordsTest.inputData="CaptureCode";
	   CF.getDataFromDataSheet(webActionKeywordsTest.inputData, "DataSheet");
	   webActionKeywordsTest.inputData = webActionKeywordsTest.dataSheetValue;
	    WF.SelectBox();
	    //Add >>
		CF.getMappedObjPropFromObjSheetByDB("Add","RowNum");
		CF.getObjectProperty_Type(webActionKeywordsTest.objectName);

		webActionKeywordsTest.ReportObjectName="PreCapture";
		WF.click();
		
		webActionKeywordsTest.inputData="2";
		WF.sleep();
		//Click Add button
		CF.getMappedObjPropFromObjSheetByDB("AddBtn","RowNum");
		CF.getObjectProperty_Type(webActionKeywordsTest.objectName);

		webActionKeywordsTest.ReportObjectName="PreCapture";
		WF.click();
		webActionKeywordsTest.inputData="2";
		WF.sleep();
		 Caller.ExecutionStatus="PASS";	      
		 Caller.Reason=webActionKeywordsTest.action+" has been Executed Successfully"; 
		 Caller.stringBuilder.append(Caller.Reason);	
		 Caller.stringBuilder.append("\n");
		 XF.AddResult(Caller.Reason, "1");
		 XF.XmlscreenShot(true);	
		 
		 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
		 Caller.StringExecutionStatus.append("\n");

		} catch (Exception e)
		{
			e.getMessage();
			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=e.toString();
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(Caller.Reason, "3");
			 XF.XmlscreenShot(true);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
		}
	}
	public  void  LogOff() throws InterruptedException, IOException
	{
		try 
		{
			//Verify
			CF.getMappedObjectPropertyFromObjectSheet("LogOff","HomePage");
			CF.getObjectProperty_Type(webActionKeywordsTest.objectName);
			webActionKeywordsTest.ReportObjectName="LogOff";
			WF.click();
			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason=webActionKeywordsTest.action+" is Executed Successfully";	       
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(Caller.Reason, "1");
			 XF.XmlscreenShot(true);
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
			 XF.XmlscreenShot(true);
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			takeSnapShot = "y";
		}
	}

	public static  void  ConfirmHandleCOR() throws InterruptedException
	{
		try 
		{
			webActionKeywordsTest.inputData="Change.*Occupation.*Code" ;
			webActionKeywordsTest.action="getHandleToWindow";
			WF.getHandleToWindow();

				WF.maximize();
				webActionKeywordsTest.inputData="2";
				WF.sleep();
				CF.getMappedObjPropFromObjSheetByDB("OKButton","RowNum");
				CF.getObjectProperty_Type(webActionKeywordsTest.objectName);
				webActionKeywordsTest.ReportObjectName="OKButton";
				webActionKeywordsTest.action="ClickButtonifExist";
				WF.ClickButtonifExist();
				webActionKeywordsTest.action="keyPress";
				WF.keyPress();
				
				webActionKeywordsTest.inputData="Change.*Occupation.*Code" ;
				webActionKeywordsTest.action="getHandleToWindowWithoutParentWindow";
				WF.getHandleToWindowWithoutParentWindow();
				webActionKeywordsTest.inputData="2";
				WF.sleep();
				CF.getMappedObjPropFromObjSheetByDB("ContinueBtn","RowNum");
			    CF.getObjectProperty_Type(webActionKeywordsTest.objectName);
					webActionKeywordsTest.ReportObjectName="ContinueBtn";
					
					webActionKeywordsTest.action="ClickButtonifExist";
					WF.ClickButtonifExist();
					webActionKeywordsTest.action="keyPress";
					WF.keyPress();
				
					webActionKeywordsTest.inputData="8";
					WF.sleep();
			 Caller.ExecutionStatus="PASS";	      
			 Caller.Reason=webActionKeywordsTest.action+" is Executed Successfully";	       
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 XF.AddResult(Caller.Reason, "1");
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
			takeSnapShot = "y";
		}
	}
	

	public static  void  AlertPopUp() throws InterruptedException
	{
		try 
		{
			WF.sleep();
			WF.sleep();
		
				CF.getMappedObjPropFromObjSheetByDB(webActionKeywordsTest.inputData,"RowNum");
				CF.getObjectProperty_Type(webActionKeywordsTest.objectName);
				webActionKeywordsTest.ReportObjectName=webActionKeywordsTest.inputData;
				WF.click();
				webActionKeywordsTest.inputData="2";
				WF.sleep();
				WebDriverWait wait = new WebDriverWait(webActionKeywordsTest.driver,10);
				
				 if(wait.until(ExpectedConditions.alertIsPresent())==null) {
					    System.out.println("alert was not present");
				 }
					else {
						Alert alert = webActionKeywordsTest.driver.switchTo().alert();
						 alert.getText();  
						 // And acknowledge the alert (equivalent to clicking "OK")
						 alert.accept();
				 }
				 // Get the text of the alert or prompt
				
			
	       Caller.ExecutionStatus="PASS";	      
			 Caller.Reason=webActionKeywordsTest.action+" is handled SuccessFully ";
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
		
		}catch (UnhandledAlertException e)
		{
			e.getMessage();
		
			 Caller.ExecutionStatus="FAIL";	      
			 Caller.Reason=e.toString();
			 Caller.stringBuilder.append(Caller.Reason);	
			 Caller.stringBuilder.append("\n");
			 
			 Caller.StringExecutionStatus.append(Caller.ExecutionStatus);	
			 Caller.StringExecutionStatus.append("\n");
			takeSnapShot = "y";
		}
	}
	
	}



