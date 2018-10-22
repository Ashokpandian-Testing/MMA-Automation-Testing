package allocator;

import java.io.File;
import java.lang.reflect.Constructor;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import businessKeywords.webActionKeywordsTest;

public class Caller {



		
		public static CommonFunctions CF = new CommonFunctions();
		public static webActionKeywordsTest WF = new webActionKeywordsTest();
		public static Xml_Reporting_Functions XF = new Xml_Reporting_Functions();
		public static String dataSheetname = "DataSheet";
		public static String runOrderSheetname = "DataSheet";
		public static String objectSheetname = "ObjectSheet";
		public static boolean loadIntoAccessDB=true;
		public static String xmlFilePath;
		public static Document document;
		public static Element TestCase;
		public static Element LOBNode,BSDNode,TCStepResult;
		
		public static String workingDir = System.getProperty("user.dir");
		
		public static String InputDir = "D:\\MMAProject\\MMA_Automation_09172018\\MMA_Automation";
	    public static String UtilitySheetpath = InputDir+"\\Execution Control\\Utility.xlsx";
	    public static String TestSummaryReportPath=InputDir+"\\Report Templates\\POS_Test Summary Report.xlsx";
	    public static String DataObjectDBpath = InputDir+"\\DataObjectDB.accdb";
		public static String BSDExecutionResultPath = InputDir+"\\Results";
		public static String OverallSummaryPath = InputDir+"\\Results";
		
		public static String IE64bitsDriverPath = InputDir+"\\IEDriverServer.exe";
       
       
		public static String projectFolderPath = workingDir;
		public static String ExecutionStatus=null;
		public static StringBuilder stringBuilder = new StringBuilder();
		public static StringBuilder StringExecutionStatus = new StringBuilder();
		public static StringBuilder BRRefernce=new StringBuilder();
		public static String BRRuleReference;
		//public static StringBuilder ScreenShotPath= new StringBuilder();
		public static String Reason=null;	
		public static Date startDate,endDate,screenshottime;
		public static long TotalTimeTaken;
		static boolean TestCaseFlag=false;
		public static String Browser;
		public static int startStepRow;
		public static int startLaunchRow;
		public static int TotalActionCount;
		public static int TotalStepCount;
		public static String Username;
		public static String Password;
		public static String LaunchSheetRow;
		public static int LOBStart;
		public static String LOBStatus;
		public static int TotalTestCaseCount;
		public static int TotalExecuted;
		public static int TotalPassed;
		public static int TotalFailed;
		public static int TotalNotExecuted;
		public static int ExecutionPercentage;
		public static int FailedPercentage;
		//public static String TestSummaryReportPath=InputDir+"\\Summary Template\\POS_Test Summary Report.xlsx";
	       public static Date LOBStartTime,LOBEndTime,CurrentTime;
		public static int TotalLOBCount,tempLObStartRow;
		
		
		public static String BSDNo,LOB,BSDTemplateFilePath;

	
		
		public static String URL;
		public static void main(String[] args) {
			
	  try
		{
		 //EnviRonemnt Setup 
		  
		  webActionKeywordsTest.UtilitySheetCount= CF.excelTotalRowCount(UtilitySheetpath,"Utility");
			for(int UtilitySheetRowNo=1;UtilitySheetRowNo <= webActionKeywordsTest.UtilitySheetCount;UtilitySheetRowNo++ )
			{
				String rowNoUtil = Integer.toString(UtilitySheetRowNo);
				String executionFlagUtil=CF.Get_Excel_ValueHeaderFirst(rowNoUtil, "Execute", "Utility", UtilitySheetpath);
		
				if(executionFlagUtil.equalsIgnoreCase(""))
				{
					break;
				}
			    if(executionFlagUtil.equalsIgnoreCase("Y"))
			      {
			  	  Browser=CF.Get_Excel_ValueHeaderFirst(rowNoUtil, "Browser", "Utility", UtilitySheetpath);
				   Username=CF.Get_Excel_ValueHeaderFirst(rowNoUtil, "Username", "Utility", UtilitySheetpath);
				   Password=CF.Get_Excel_ValueHeaderFirst(rowNoUtil, "Password", "Utility", UtilitySheetpath);
			       webActionKeywordsTest.browser = Browser;
			       CF.close_browser(Browser);
			    
			     URL=CF.Get_Excel_ValueHeaderFirst(rowNoUtil, "URL", "Utility", UtilitySheetpath);
			
		         break;
			  }
			   
			}
		
	
			//WF.CreatePOSDashboardFile();
			
			//test.copySheets(WF.BSDFinalReportPath);
			//CF.copyResultFormat(WF.BSDFinalReportPath);
			
    //BSD Level Iteration
			int LaunchSheetRowCount= CF.excelTotalRowCount(UtilitySheetpath,"Launch Sheet");
			startLaunchRow=1;
			BSDLoop:
			for(int LaunchSheetRow=startLaunchRow;LaunchSheetRow <= LaunchSheetRowCount; )
				
			{
				String rowNoString = Integer.toString(LaunchSheetRow);
				String executionFlag=CF.Get_Excel_ValueHeaderFirst(rowNoString, "BSD Execution Flag", "Launch Sheet", UtilitySheetpath);
				if(executionFlag.equalsIgnoreCase(""))
				{
					break;
				}
				if(executionFlag.equalsIgnoreCase("Y")) {
					
					BSDNo=CF.Get_Excel_ValueHeaderFirst(rowNoString, "BSD No", "Launch Sheet", UtilitySheetpath);
					BSDTemplateFilePath=CF.Get_Excel_ValueHeaderFirst(rowNoString, "BSD Template File Path", "Launch Sheet", UtilitySheetpath);				
					
				    TotalLOBCount=WF.getNbOfMergedRegions(LaunchSheetRow,"A");
				    if(BSDNo.equalsIgnoreCase("Smoke Testing"))
					{
				    	 WF.CreateSmokeTestingResultExcelFile();
						  WF.CreateSmokeDashboardFile();	
						  XF.CreateCustomXmlNew();
					}
				    else
				    {
				    //Copy Test Summary Report
					WF.CreateSummaryReport();
				    WF.CreateExcelFile();
				    WF.CreatePOSDashboardFile();
				    XF.CreateCustomXmlNew();
				    }
					
				    //LOB Level Iteration
					 tempLObStartRow= LaunchSheetRow;
				   for(int LObStartRow=1;LObStartRow <= TotalLOBCount; LObStartRow++) {
				
					String rowNoLOB = Integer.toString(tempLObStartRow);
					String executionLOBFlag=CF.Get_Excel_ValueHeaderFirst(rowNoLOB, "LOB Execution Flag" , "Launch Sheet", UtilitySheetpath);
					LOB=CF.Get_Excel_ValueHeaderFirst(rowNoLOB, "LOB" , "Launch Sheet", UtilitySheetpath);
					if(executionLOBFlag.equalsIgnoreCase(""))
					{
						break;
					}
					if(executionLOBFlag.equalsIgnoreCase("Y")) {
						Calendar calendar = Calendar.getInstance();
					    LOBStartTime = calendar.getTime();					    
					    XF.AddLOBNode();
						LOBStart = 0;
						TotalTestCaseCount = 0;
						TotalExecuted = 0;
						TotalPassed = 0;
						TotalFailed = 0;
						TotalNotExecuted = 0;
						ExecutionPercentage = 0;
					
						// TODO Auto-generated method stub
						int runOrderSheetRowCount= CF.excelTotalRowCount(BSDTemplateFilePath,"Automation Test Case");
						webActionKeywordsTest.objectSheetRowCount= CF.excelTotalRowCount(BSDTemplateFilePath,Caller.objectSheetname);
						if(loadIntoAccessDB == true)
						{
							
							CF.executeNonQueryInAccessDB("delete * from Objects");

						
							for(int eachRowObjSheet = 1;eachRowObjSheet <= webActionKeywordsTest.objectSheetRowCount; eachRowObjSheet++)
							{
								String rowNoRunorder = Integer.toString(eachRowObjSheet);
								String eachRowObjname = CF.Get_Excel_ValueHeaderFirst(rowNoRunorder, "ObjectName", Caller.objectSheetname, Caller.BSDTemplateFilePath);
								if(!eachRowObjname.equalsIgnoreCase(""))
								{
									boolean queryExecuted = CF.executeNonQueryInAccessDB("insert into Objects (RowNum,ObjectName) values (" + eachRowObjSheet + ",'" + eachRowObjname + "')");
									if(queryExecuted == false)
									{
										System.out.println("Error in loading the data of Object Sheet in row Number : "+ eachRowObjSheet);
										//DBloadingErrorStatus= true;
									}
								}
							}
						     }
						//webActionKeywordsTest.AutomationStepCount= CF.excelTotalRowCount(BSDTemplateFilePath,"Automation Test Case");
						webActionKeywordsTest.dataSheetRowCount= CF.excelTotalRowCount(BSDTemplateFilePath,"DataSheet");
						startStepRow=4;
						//Test Case Level Iteration
						Loop:
						for(int runOrderSheetRowNo=startStepRow;runOrderSheetRowNo <= runOrderSheetRowCount; )
						{
							 TestCaseFlag=false;
							String rowNoRunorder = Integer.toString(runOrderSheetRowNo);
							String executionTestFlag=CF.Get_Excel_Value(rowNoRunorder, "Execution Flag", "Automation Test Case", BSDTemplateFilePath);
							webActionKeywordsTest.testCaseID=CF.Get_Excel_Value(rowNoRunorder, "Test Case ID", "Automation Test Case", BSDTemplateFilePath);
							TotalStepCount=CF.findRowInput(startStepRow, "LogOff",12, runOrderSheetRowCount, "Automation Test Case",BSDTemplateFilePath);
							TotalTestCaseCount = TotalTestCaseCount+1;
							
							
							if(webActionKeywordsTest.testCaseID.equalsIgnoreCase(""))
							{
								break;
							}
							
							if(executionTestFlag.equalsIgnoreCase("Y"))
							{
								 TestCaseFlag=true;
								 XF.AddTestCaseNode();
									 stringBuilder.setLength(0);
								     StringExecutionStatus.setLength(0);
									 BRRefernce.setLength(0);
								    webActionKeywordsTest.driver = CF.launchBrowser(Browser);
									webActionKeywordsTest.driver.get(URL);
									webActionKeywordsTest.driver.manage().window().maximize();
									TotalExecuted=TotalExecuted+1;
								     webActionKeywordsTest.dataSheetTestCaseRowIndex = CF.findRowIndexLOB(webActionKeywordsTest.testCaseID, "TestCaseName","LOB", webActionKeywordsTest.dataSheetRowCount, "DataSheet",BSDTemplateFilePath);
								//Comment have to give different path for utility
								
								//Test Action Step Level Iteration
									//TotalStepCount=CF.findRowInput(startStepRow, "LogOff",10, runOrderSheetRowCount, "Automation Test Case",BSDTemplateFilePath);
									//TotalActionCount=CF.findRowIndex("LogOff",9,runOrderSheetRowCount,"Automation Test Case",BSDTemplateFilePath);
									for(int testcaseSheetRowNo =0; testcaseSheetRowNo <= TotalStepCount; testcaseSheetRowNo++)
									{
										int StepIter;
										 StepIter=runOrderSheetRowNo;
										String testcaserowNoString = Integer.toString(StepIter);
										
									    webActionKeywordsTest.action = CF.Get_Excel_Action_Value(StepIter, 12, "Automation Test Case", BSDTemplateFilePath);
										BRRuleReference= CF.Get_Excel_Action_Value(StepIter, 5, "Automation Test Case", BSDTemplateFilePath);
										if(BRRuleReference.contains("BR")) {
											 Caller.BRRefernce.append(BRRuleReference);	
											 Caller.BRRefernce.append("\n");
											 
										}
										if(webActionKeywordsTest.action.equalsIgnoreCase(""))
										{
											runOrderSheetRowNo=StepIter+1;
											continue ;
											
										}
										//webActionKeywordsTest.actionRowNo = CF.Get_Excel_Value(testcaserowNoString, "Step", webActionKeywordsTest.testFlowSheetName, BSDTemplateFilePath);
										//webActionKeywordsTest.PageName = CF.Get_Excel_Value(testcaserowNoString, "Page Name", "Automation Test Case", BSDTemplateFilePath);
									/*	webActionKeywordsTest.PageName=   webActionKeywordsTest.driver.getTitle();
										if(webActionKeywordsTest.PageName.contains("/")) {
											webActionKeywordsTest.PageName.replace("/","_");
										}*/
										webActionKeywordsTest.objectName = CF.Get_Excel_Value(testcaserowNoString, "ObjectName", "Automation Test Case", BSDTemplateFilePath);
										webActionKeywordsTest.ReportObjectName=webActionKeywordsTest.objectName;
										
										if(!webActionKeywordsTest.objectName.equalsIgnoreCase(""))
										{
											CF.getMappedObjPropFromObjSheetByDB(webActionKeywordsTest.objectName,"RowNum");
											CF.getObjectProperty_Type(webActionKeywordsTest.objectName);
										}
										

										webActionKeywordsTest.inputData = CF.Get_Excel_Value(testcaserowNoString, "Utility Data", "Automation Test Case", BSDTemplateFilePath);
										
										Boolean getTestDataFlag = false;
										
									
										if(!webActionKeywordsTest.inputData.contains("~") && !webActionKeywordsTest.inputData.contains(";") && !webActionKeywordsTest.inputData.contains("'")&& !webActionKeywordsTest.inputData.contains(" ")&& !webActionKeywordsTest.inputData.isEmpty())
										{
											
											if(WF.isNumeric()){
												
												getTestDataFlag = false;
											}
											
											else {
												
												if (webActionKeywordsTest.inputData.contains("Map")) {
													webActionKeywordsTest.inputData=webActionKeywordsTest.inputData;
												}
												
												else {
												getTestDataFlag = CF.getDataFromDataSheet(webActionKeywordsTest.inputData,"DataSheet");	
											}
											}
											
											if(getTestDataFlag == true)
											{
												webActionKeywordsTest.inputData = webActionKeywordsTest.dataSheetValue;
											}
											
										}
										else {
											getTestDataFlag = false;
										}
							
										if(webActionKeywordsTest.action.equalsIgnoreCase("LogOff"))
										{
											startStepRow=runOrderSheetRowNo+1;
											runOrderSheetRowNo=startStepRow;
											invokeBusinessComponent(webActionKeywordsTest.action);
											if(BSDNo.equalsIgnoreCase("Smoke Testing"))
											{
											 WF.ExcelResultwriteSmoke(webActionKeywordsTest.BSDFinalResultPath, "Result Sheet");
											 XF.AddTestCaseStatusAtt();
											}
										    else
										    {
										
										    	WF.ExcelResultwrite(webActionKeywordsTest.BSDFinalResultPath, "Result Sheet");
										    	 XF.AddTestCaseStatusAtt();
										    }
										 
										
										
											  CF.close_browser(webActionKeywordsTest.browser);
									         continue  Loop;
										}
								
										if(!webActionKeywordsTest.action.equalsIgnoreCase(""))
										{
											invokeBusinessComponent(webActionKeywordsTest.action);
											if(Caller.ExecutionStatus.contains("FAIL")){
												 if(BSDNo.equalsIgnoreCase("Smoke Testing"))
													{
													 WF.ExcelResultwriteSmoke(webActionKeywordsTest.BSDFinalResultPath, "Result Sheet");
													 XF.AddTestCaseStatusAtt();
													}
												    else
												    {
												
												    	WF.ExcelResultwrite(webActionKeywordsTest.BSDFinalResultPath, "Result Sheet");
												    	 XF.AddTestCaseStatusAtt();
												    }
											
												startStepRow=TotalStepCount+startStepRow;
												runOrderSheetRowNo=startStepRow;
												CF.close_browser(webActionKeywordsTest.browser);
											     
												continue Loop;
											}
											runOrderSheetRowNo++;
										}
										
										
//										BRRuleReference = null;	
								
							}
									
						}
							if(TestCaseFlag==true) {
								if(BSDNo.equalsIgnoreCase("Smoke Testing"))
								{
								 WF.ExcelResultwriteSmoke(webActionKeywordsTest.BSDFinalResultPath, "Result Sheet");
								 XF.AddTestCaseStatusAtt();
								}
							    else
							    {
							
							    	WF.ExcelResultwrite(webActionKeywordsTest.BSDFinalResultPath, "Result Sheet");
							    	 XF.AddTestCaseStatusAtt();
							    }
						
						}
							startStepRow=TotalStepCount+startStepRow;
							runOrderSheetRowNo=startStepRow;
						}
						
			
					//Reporting
					     CF.close_browser(webActionKeywordsTest.browser);
					     
					     if (!(Caller.TotalFailed > 0)){
						     LOBStatus = "PASSED";
					   }
						   else
						   {
							   LOBStatus = "FAILED";
						   }
				
					   TotalNotExecuted = TotalTestCaseCount - TotalExecuted;
				
					   ExecutionPercentage = TotalPassed/TotalExecuted*100;
					   
					   FailedPercentage = TotalFailed/TotalExecuted*100;
			
					   WF.DateCalculation();
					   if(BSDNo.equalsIgnoreCase("Smoke Testing"))
						{
						   WF.ExcelSmokeDashboardwrite(webActionKeywordsTest.BSDFinalResultPath, webActionKeywordsTest.POSSheetName);
						   XF.AddLobSummaryXml();
						}
					    else
					    {
					
					    	WF.ExcelPOSDashboardwrite(webActionKeywordsTest.BSDFinalReportPath, webActionKeywordsTest.POSSheetName);
					    	 XF.AddLobSummaryXml();
					    }
					   
					     
						}
					tempLObStartRow=tempLObStartRow+1;
						}
					
					
					
					tempLObStartRow=tempLObStartRow+1;

					}
					
					

				else {
					Thread.sleep(1000);
					TotalLOBCount=WF.getNbOfMergedRegions(LaunchSheetRow,"A");
					 startLaunchRow=LaunchSheetRow+TotalLOBCount;
					 LaunchSheetRow= startLaunchRow;
					 continue BSDLoop;
				}
				 startLaunchRow=tempLObStartRow;
				 LaunchSheetRow= startLaunchRow;
				 XF.AddLOBEndTime();
			
					}	
		
		}

			catch (Exception e)
		{
			e.getMessage();
			 if(BSDNo.equalsIgnoreCase("Smoke Testing"))
				{
				 WF.ExcelResultwriteSmoke(webActionKeywordsTest.BSDFinalResultPath, "Result Sheet");
				 XF.AddTestCaseStatusAtt();
				}
			    else
			    {
			
			    	WF.ExcelResultwrite(webActionKeywordsTest.BSDFinalResultPath, "Result Sheet");
			    	 XF.AddTestCaseStatusAtt();
			    }
			System.out.println(e);
	
		}
	}

			
	
	
	/**
	 * Function to invoke the business component corresponding to the keyword passed
	 * @param currentKeyword The keyword representing the business component
	 */
	public static void invokeBusinessComponent(String currentKeyword)
	{
		Boolean isMethodFound = false;
		final String CLASS_FILE_EXTENSION = ".class";
		File packageDirectory = new File(Caller.projectFolderPath + "\\bin\\businessKeywords");
		File[] packageFiles = packageDirectory.listFiles();

		for (int i = 0; i < packageFiles.length; i++)
		{
			File packageFile = packageFiles[i];
			String fileName = packageFile.getName();

			//We only want the .class files
			if (fileName.endsWith(CLASS_FILE_EXTENSION))
			{
				//Remove the .class extension to get the class name
				String className = fileName.substring(0, fileName.length() - CLASS_FILE_EXTENSION.length());
				try
				{
					Class<?> reusableComponents = Class.forName("businessKeywords." + className);
					Method executeComponent;

					try {
						//convert the first letter of the method to lowercase (in line with java naming conventions)
						//currentKeyword = currentKeyword.substring(0, 1).toLowerCase() + currentKeyword.substring(1);
						executeComponent = reusableComponents.getMethod(currentKeyword, (Class<?>[]) null);
						
					} catch(NoSuchMethodException ex) {
						//If the method is not found in this class, search the next class
						continue;
					}

					isMethodFound = true;

					Constructor<?> ctor = reusableComponents.getDeclaredConstructors()[0];
					//Object businessComponent = ctor.newInstance(scriptHelper, driver);
					Object businessComponent = ctor.newInstance();

					executeComponent.invoke(businessComponent, (Object[]) null);

					break;
				} catch (Exception ex) {

				}
			} 
		}
	}

	






  }

