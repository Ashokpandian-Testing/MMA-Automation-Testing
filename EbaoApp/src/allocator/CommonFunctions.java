package allocator;

import java.awt.AWTException;
import java.awt.Robot;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;
import java.util.NoSuchElementException;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;


import businessKeywords.webActionKeywordsTest;
import jxl.Workbook;



public class CommonFunctions {

	public static WebDriver launchBrowser(String browser)
	{
		WebDriver driver = null;

		if (browser.equalsIgnoreCase("Firefox"))
		{
			//System.setProperty("webdriver.firefox.bin",EnvironmentVariable.firefoxBrowserPath);
			driver = new FirefoxDriver();
		}
		else if (browser.equalsIgnoreCase("IEx86"))
		{
			//System.setProperty("webdriver.iexplore.bin", EnvironmentVariable.IE32bitsDriverPath);
			//System.setProperty("webdriver.ie.driver", EnvironmentVariable.IE32bitsDriverPath);
			DesiredCapabilities capabilities = new DesiredCapabilities();
			capabilities.setCapability("ignoreProtectedModeSettings", true);
			driver = new InternetExplorerDriver(capabilities);
		}
		else if (browser.equalsIgnoreCase("IEx64"))
		{
			System.setProperty("webdriver.ie.driver", Caller.IE64bitsDriverPath);
			//System.setProperty("webdriver.iexplore.bin", EnvironmentVariable.IE64bitsDriverPath);
			DesiredCapabilities capabilities = new DesiredCapabilities();
			//capabilities.setCapability(InternetExplorerDriver.INITIAL_BROWSER_URL, "http://192.168.88.23:7001/ls/loginPage.do");
			
			capabilities.setCapability("ignoreProtectedModeSettings", true);
			driver = new InternetExplorerDriver(capabilities);
		}
		else if (browser.equalsIgnoreCase("Chromex86"))
		{

			//System.setProperty("webdriver.chrome.driver",EnvironmentVariable.chrome32bitsDriverPath);
			DesiredCapabilities capabilities = new DesiredCapabilities();
			capabilities.setCapability("ignoreProtectedModeSettings", true);
			driver = new ChromeDriver(capabilities);
		}
		else if (browser.equalsIgnoreCase("Chromex64"))
		{

			//System.setProperty("webdriver.chrome.driver",EnvironmentVariable.chrome64bitsDriverPath);
			DesiredCapabilities capabilities = new DesiredCapabilities();
			capabilities.setCapability("ignoreProtectedModeSettings", true);
			driver = new ChromeDriver(capabilities);
		}

		return driver;
	}
	
	
	//comom


	
	public static int excelTotalRowCount(String sheetpath, String sheetName)
	{
		try
		{
			//HSSFWorkbook book = null;
			org.apache.poi.ss.usermodel.Workbook book = null;
			File file=new File(sheetpath);   
            if(file.exists())   
            {   
            	InputStream myxls = new FileInputStream(sheetpath);
            	if(!(sheetpath.contains(".xlsx")))
            	{
            		book = new HSSFWorkbook(myxls);
            	}else
            	{
            		//book = (org.apache.poi.ss.usermodel.Workbook) new HSSFWorkbook(new POIFSFileSystem(myxls));
            		//book = (org.apache.poi.ss.usermodel.Workbook) new XSSFWorkbook(myxls);
            		book = new XSSFWorkbook(myxls);
            	}
            	//HSSFSheet sheet = book.getSheetAt(0);    
            	//HSSFSheet sheet = book.getSheet(sheetName);
            	Sheet sheet = book.getSheet(sheetName);
            	System.out.println("Sheet row count.." +sheet.getLastRowNum());
            	int rowCount = sheet.getLastRowNum();
            	myxls.close();
            	return rowCount;

            } 
			
		} catch(Exception e)
		{
			e.getMessage();
            System.out.println(e);
		}
		return -1;
	}
	
	
	
	
	public static int excelTotalColCount(String sheetpath, String sheetName)
	{
		int colCount = -1;
		try
		{
			Row row = null;   
			org.apache.poi.ss.usermodel.Workbook book = null;
			File file=new File(sheetpath);   
            if(file.exists())   
            {   
            	InputStream myxls = new FileInputStream(sheetpath);
            	if(!(sheetpath.contains(".xlsx")))
            	{
            		book = new HSSFWorkbook(myxls);
            	}else
            	{
            		//book = (org.apache.poi.ss.usermodel.Workbook) new HSSFWorkbook(new POIFSFileSystem(myxls));
            		//book = (org.apache.poi.ss.usermodel.Workbook) new XSSFWorkbook(myxls);
            		book = new XSSFWorkbook(myxls);
            	}
            	Sheet sheet = book.getSheet(sheetName);
            	row=((org.apache.poi.ss.usermodel.Sheet) sheet).getRow(0);
		        if(!(row== null)){
		              colCount=row.getLastCellNum();
		              System.out.println("Column Count:"+colCount);
		        }
		        else
		        {                     
		              colCount=0;       
		        }
		        
            	
            	myxls.close();
            	return colCount;

            } 
			
		} catch(Exception e)
		{
			e.getMessage();
            System.out.println(e);
		}
		return -1;
	}
	
	
	
		
		//********* Finding the particular row index in the sheet *****************//
		public static int findRowIndex(String row_name, String Column_Name, int totalRowCount, String Sheet_Name,String Input_file) 
		{
			String value="";
			int Row_Num = -1;   
			try 
			{
				org.apache.poi.ss.usermodel.Workbook tempWB = null;               
				// String WB_Name = "C:\\ACES\\ACES_Claimvalidation\\com\\WLP\\AutomationComponents\\Application\\Data\\Input\\TestData.xls";
				String WB_Name = Input_file;

				//System.out.println("Excel Filepath:"+ WB_Name);
				InputStream inp =null;              
				try 
				{
					inp = new FileInputStream(WB_Name);
					if(!(WB_Name.contains(".xlsx")))
					{
						tempWB = new HSSFWorkbook(inp);
						//System.out.println(tempWB);
					}else
					{
						//tempWB = (org.apache.poi.ss.usermodel.Workbook) new HSSFWorkbook(new POIFSFileSystem(inp));
						tempWB = new XSSFWorkbook(inp);
					}

					Sheet s1 = tempWB.getSheet(Sheet_Name);
					int Col_Num =findColNew(s1,Column_Name);

					for(int excelRowNo = 0; excelRowNo<=totalRowCount;excelRowNo++)
					{
						Row r = s1.getRow(excelRowNo);
						Cell  cel = r.getCell(Col_Num);
						switch (cel.getCellType()) {
	                    case Cell.CELL_TYPE_STRING:
	                    	value = cel.getStringCellValue();
	                     
	                        break;
	                    case Cell.CELL_TYPE_BOOLEAN:
	                    	boolean flag = cel.getBooleanCellValue();
	                    	value=String.valueOf(flag);
	                     
	                        break;
	                    case Cell.CELL_TYPE_NUMERIC:
	                    double number = cel.getNumericCellValue();
	                    	value=String.valueOf(number).trim();
	                        break;
	                    case Cell.CELL_TYPE_BLANK:
	                    	value="";
		                        break;
						}
						
						if(value.equalsIgnoreCase(row_name))
						{
							Row_Num = excelRowNo;
							break;
						}
						
					}
					
					inp.close();
					return Row_Num; 
					
				} catch (IOException e) {
					System.out.println(e);

					e.getMessage();
				} catch (Exception e) {

					e.getMessage();
					System.out.println(e);
				}

			} catch (Exception e) {
				System.out.println(" Excel_Read - Function issue : "+e);
				e.getMessage();
			}
			return Row_Num;  
		}
		
		
		
		
		//********* Finding the particular row index in the sheet *****************//
				public static int findRowIndex(String row_name, int Column_Index, int totalRowCount, String Sheet_Name,String Input_file) 
				{
					String value="";
					int Row_Num = -1;   
					try 
					{
						org.apache.poi.ss.usermodel.Workbook tempWB = null;               
						// String WB_Name = "C:\\ACES\\ACES_Claimvalidation\\com\\WLP\\AutomationComponents\\Application\\Data\\Input\\TestData.xls";
						String WB_Name = Input_file;

						//System.out.println("Excel Filepath:"+ WB_Name);
						InputStream inp =null;              
						try 
						{
							inp = new FileInputStream(WB_Name);
							if(!(WB_Name.contains(".xlsx")))
							{
								tempWB = new HSSFWorkbook(inp);
								//System.out.println(tempWB);
							}else
							{
								//tempWB = (org.apache.poi.ss.usermodel.Workbook) new HSSFWorkbook(new POIFSFileSystem(inp));
								tempWB = new XSSFWorkbook(inp);
							}

							Sheet s1 = tempWB.getSheet(Sheet_Name);
							int Col_Num = Column_Index;
							int testcasestartactiontemp=0;
							for(int excelRowNo = 0; excelRowNo<totalRowCount;excelRowNo++)
							{
							
								Row r = s1.getRow(excelRowNo);
								Cell  cel = r.getCell(Col_Num);
								switch (cel.getCellType()) {
			                    case Cell.CELL_TYPE_STRING:
			                    	value = cel.getStringCellValue();
			                    	testcasestartactiontemp++;
			                        break;
			                    case Cell.CELL_TYPE_BOOLEAN:
			                    	boolean flag = cel.getBooleanCellValue();
			                    	value=String.valueOf(flag);
			                    	testcasestartactiontemp++;
			                        break;
			                    case Cell.CELL_TYPE_NUMERIC:
			                    double number = cel.getNumericCellValue();
			                    	value=String.valueOf(number).trim();
			                    	testcasestartactiontemp++;
			                        break;
			                    case Cell.CELL_TYPE_BLANK:
			                    	value="";
			                    	testcasestartactiontemp++;
				                        break;
								}
							
							
								if(value.equalsIgnoreCase(row_name))
								{
									Row_Num = testcasestartactiontemp;
									break;
								}
								
							}
							
							inp.close();
							return Row_Num; 
							
						} catch (IOException e) {
							System.out.println(e);

							e.getMessage();
						} catch (Exception e) {

							e.getMessage();
							System.out.println(e);
						}

					} catch (Exception e) {
						System.out.println(" Excel_Read - Function issue : "+e);
						e.getMessage();
					}
					return Row_Num;  
				}
				
				
				
		
		
		//********* Finding the particular column in the sheet *****************//
				public static int findCol(Sheet sheet, String colName) {
			        Row row = null;         
			        int colCount=0;

			        row=((org.apache.poi.ss.usermodel.Sheet) sheet).getRow(3);
			        if(!(row== null)){
			              colCount=row.getLastCellNum();
			           
			        }
			        else{                     
			              colCount=0;       
			        }
			        for(int j=0;j<colCount;j++){
			              if(!( row.getCell(j)==null)){
//			            	  String sample=row.getCell(j).toString();
//			            	  System.out.println(sample);
			                    if(row.getCell(j).toString().trim().equalsIgnoreCase(colName) || row.getCell(j).toString().trim().equalsIgnoreCase((colName+"[][String]"))){
			                      
			                        return j;
			                    }
			              }
			        }
//			       System.out.println(colName + " ----  column is not present in excel sheet.");
			        //logWarning(colName + " column is not present in excel sheet.");
			        return -1;  
			  }
				
		
				/********* Finding the particular column in the sheet *****************/
				public static int findColNew(Sheet sheet, String colName) {
			        Row row = null;         
			        int colCount=0;

			        row=((org.apache.poi.ss.usermodel.Sheet) sheet).getRow(0);
			        if(!(row== null)){
			              colCount=row.getLastCellNum();
			         
			        }
			        else{                     
			              colCount=0;       
			        }
			        for(int j=0;j<colCount;j++){
			              if(!( row.getCell(j)==null)){
//			            	  String sample=row.getCell(j).toString();
//			            	  System.out.println(sample);
			                    if(row.getCell(j).toString().trim().equalsIgnoreCase(colName) || row.getCell(j).toString().trim().equalsIgnoreCase((colName+"[][String]"))){
			                
			                        return j;
			                    }
			              }
			        }
//			       System.out.println(colName + " ----  column is not present in excel sheet.");
			        //logWarning(colName + " column is not present in excel sheet.");
			        return -1;  
			  }
				
		
		
		
		//********* Getting the  values from a sheet *****************//
		
		public String Get_Excel_Value(String row_num, String Column_Name,String Sheet_Name,String Input_file)
	    {           
	      String value="";
	      
	          try {
	                
	                int Row_Num = 0;              
	                Row_Num = Integer.parseInt(row_num);

	                org.apache.poi.ss.usermodel.Workbook tempWB = null;               
	               // String WB_Name = "C:\\ACES\\ACES_Claimvalidation\\com\\WLP\\AutomationComponents\\Application\\Data\\Input\\TestData.xls";
	                String WB_Name = Input_file;
	                
	                //System.out.println("Excel Filepath:"+ WB_Name);
	                InputStream inp =null;              
	                try {
	                      inp = new FileInputStream(WB_Name);
	                      if(!(WB_Name.contains(".xlsx")))
	                      {
	                    	  tempWB = new HSSFWorkbook(inp);
	                    
	                      }else
	                      {
	                    	  //tempWB = (org.apache.poi.ss.usermodel.Workbook) new HSSFWorkbook(new POIFSFileSystem(inp));
	                    	  tempWB = new XSSFWorkbook(inp);
	                      }

	                      Sheet s1 = tempWB.getSheet(Sheet_Name);
	                      int Col_Num =findCol(s1,Column_Name);
	                      Row r = s1.getRow(Row_Num);
	                      Cell  cel = r.getCell(Col_Num);
	                     
	              		switch (cel.getCellType()) {
	                    case Cell.CELL_TYPE_STRING:
	                    	value = cel.getStringCellValue();
	                   
	                        break;
	                    case Cell.CELL_TYPE_BOOLEAN:
	                    	boolean flag = cel.getBooleanCellValue();
	                    	value=String.valueOf(flag);
	                
	                        break;
	                    case Cell.CELL_TYPE_NUMERIC:
	                    double number = cel.getNumericCellValue();
	                    	value=String.valueOf(number).trim();
	              
	                        break;
	                    case Cell.CELL_TYPE_BLANK:
	                    	value="";
	                    
		                        break;
	              		}
	                      inp.close();
	                      
	                } catch (IOException e) {
	                  System.out.println(e);
	                     
	                      e.getMessage();
	                } catch (Exception e) {
	                      
	                      e.getMessage();
	                     System.out.println(e);
	                }
	                          
	          } catch (Exception e) {
	                System.out.println(" Excel_Read - Function issue : "+e);
	                e.getMessage();
	          }
	            return value;
	                        
	    }
		
		
		
		
		//********* Getting the  values from a sheet *****************//
		
				public String Get_Excel_ValueHeaderFirst(String row_num, String Column_Name,String Sheet_Name,String Input_file)
			    {           
			      String value="";
			      
			          try {
			                
			                int Row_Num = 0;              
			                Row_Num = Integer.parseInt(row_num);

			                org.apache.poi.ss.usermodel.Workbook tempWB = null;               
			               // String WB_Name = "C:\\ACES\\ACES_Claimvalidation\\com\\WLP\\AutomationComponents\\Application\\Data\\Input\\TestData.xls";
			                String WB_Name = Input_file;
			                
			                //System.out.println("Excel Filepath:"+ WB_Name);
			                InputStream inp =null;              
			                try {
			                      inp = new FileInputStream(WB_Name);
			                      if(!(WB_Name.contains(".xlsx")))
			                      {
			                    	  tempWB = new HSSFWorkbook(inp);
			                    	 
			                      }else
			                      {
			                    	  //tempWB = (org.apache.poi.ss.usermodel.Workbook) new HSSFWorkbook(new POIFSFileSystem(inp));
			                    	  tempWB = new XSSFWorkbook(inp);
			                      }

			                      Sheet s1 = tempWB.getSheet(Sheet_Name);
			                      int Col_Num =findColNew(s1,Column_Name);
			                      Row r = s1.getRow(Row_Num);
			                      Cell  cel = r.getCell(Col_Num);
			                      
				              		switch (cel.getCellType()) {
				                    case Cell.CELL_TYPE_STRING:
				                    	value = cel.getStringCellValue();
				                   
				                        break;
				                    case Cell.CELL_TYPE_BOOLEAN:
				                    	boolean flag = cel.getBooleanCellValue();
				                    	value=String.valueOf(flag);
				                
				                        break;
				                    case Cell.CELL_TYPE_NUMERIC:
				                    double number = cel.getNumericCellValue();
				                    	value=String.valueOf(number).trim();
				              
				                        break;
				                    case Cell.CELL_TYPE_BLANK:
				                    	value="";
				                    
					                        break;
				              		}
			                  
			                      inp.close();
			                      
			                } catch (IOException e) {
			                  System.out.println(e);
			                     
			                      e.getMessage();
			                } catch (Exception e) {
			                      
			                      e.getMessage();
			                     System.out.println(e);
			                }
			                          
			          } catch (Exception e) {
			                System.out.println(" Excel_Read - Function issue : "+e);
			                e.getMessage();
			          }
			            return value;
			                        
			    }
				
				
							//System.out.println("Excel Filepath:"+ WB_Name);
						 
		 public void createFolder(String newFolderPath)
		 {
			 try
			 {
				 File file = new File(newFolderPath);
				 if (!file.exists()) 
				 {
					 if (file.mkdir()) {
						 System.out.println("Directory is created!");
					 } else 
					 {
						 System.out.println("Failed to create directory!");
					 }
				 }
			 }catch (Exception e)
			 {
				e.getMessage();
				System.out.println(e);
			 }
		 }
		 
		 

		 public void copyPasteFile(String copyFilePath, String pasteFilePath)
		 {
			 try
			 {
				 File source = new File(copyFilePath); 
			
					File dest = new File(pasteFilePath); 
					long start = System.nanoTime(); 
					try 
					{
						copyFileUsingStream(source, dest);
					} catch (IOException e2) {
						e2.printStackTrace();
					} 
					System.out.println("Time taken by Stream Copy = "+(System.nanoTime()-start)); 

				 
			 }catch (Exception e)
			 {
				e.getMessage();
				System.out.println(e);
			 }
		 }
		 
		 
		 private static void copyFileUsingStream(File source, File dest) throws IOException { 
				InputStream is = null; 
				OutputStream os = null; 
				try { 
					is = new FileInputStream(source); 
					os = new FileOutputStream(dest); 
					byte[] buffer = new byte[1024]; 
					int length; 
					while ((length = is.read(buffer)) > 0) 
					{ 
						os.write(buffer, 0, length); 
					} 
				} finally { 
					is.close(); 
					os.close(); 
				} 
			}

		 

		 
		 
		 
		 public  void  close_browser(String browser) throws InterruptedException
			{
				
				try 
				{
					browser = browser.toLowerCase();
					if(browser.startsWith("firefox"))
					{
						Runtime.getRuntime().exec("taskkill /F /IM firefox.exe");
						System.out.println("FIREFOX browser is closed successfully");
						Thread.sleep(500);
					}else if(browser.startsWith("ie"))
					{
						Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe");
						System.out.println("IE browser is closed successfully");
						Thread.sleep(500);
					}else if(browser.startsWith("chrome"))
					{
						Runtime.getRuntime().exec("taskkill /F /IM chrome.exe");
						Runtime.getRuntime().exec("taskkill /F /IM chromedriverx32.exe");
						System.out.println("CHROME browser is closed successfully");
						Thread.sleep(500);
					}
					
				} catch (IOException e1) {
					
					e1.printStackTrace();
				} catch (Exception e)
				{
					e.getMessage();
					System.out.println(e);
				}
			}


		


		 public boolean isElementPresent(By by) 
		 {
			 try {
				 //driver.findElement(by);
				 if(webActionKeywordsTest.driver.findElement(by) != null)
				 {
					 return true;
				 }else
				 {
					 return false;
				 }

			 } catch (NoSuchElementException e) 
			 {
				 return false;
			 } catch (Exception e)
			 {
				 return false;
			 }
		 }
		 
		 
				


		
		 
		 
		 
		 
		 public  void  getObjectProperty_Type(String objectName) throws InterruptedException
		 {
			 try 
			 {
				webActionKeywordsTest.objectProperty = "";
				webActionKeywordsTest.objectType = "";
				
				if(objectName.startsWith("css="))
				{
					webActionKeywordsTest.objectProperty = objectName.split("css=")[1].trim();
					webActionKeywordsTest.objectType = "css";
				}else if(objectName.startsWith("link="))
				{
					webActionKeywordsTest.objectProperty = objectName.split("link=")[1].trim();
					webActionKeywordsTest.objectType = "link";
				}else if(objectName.startsWith("id="))
				{
					webActionKeywordsTest.objectProperty = objectName.split("id=")[1].trim();
					webActionKeywordsTest.objectType = "id";
				}else if(objectName.startsWith("name="))
				{
					webActionKeywordsTest.objectProperty = objectName.split("name=")[1].trim();
					webActionKeywordsTest.objectType = "name";
				}else if(objectName.startsWith("xpath="))
				{
					webActionKeywordsTest.objectProperty = objectName.split("xpath=")[1].trim();
					webActionKeywordsTest.objectType = "xpath";
				}else if(objectName.startsWith("//") || objectName.startsWith("(//"))
				{
					webActionKeywordsTest.objectProperty = objectName;
					webActionKeywordsTest.objectType = "xpath";
				}
			 
			 }catch (Exception e)
			 {
				 e.getMessage();
				 System.out.println(e);
			 }
		 }

		 
		 public  void  getMappedObjectPropertyFromObjectSheet(String objectName,String PageName) throws InterruptedException
		 {
			 try 
			 {
				 for(int objShtRowNo=1; objShtRowNo <= webActionKeywordsTest.objectSheetRowCount; objShtRowNo++)
				 {
					 String objShtRowNoString = Integer.toString(objShtRowNo);
					 String object_Name=Get_Excel_ValueHeaderFirst(objShtRowNoString, "ObjectName", Caller.objectSheetname, Caller.BSDTemplateFilePath);
					
					 if(objectName.equalsIgnoreCase(object_Name))
					 {
						 webActionKeywordsTest.objectName =Get_Excel_ValueHeaderFirst(objShtRowNoString, "ObjectProperty", Caller.objectSheetname, Caller.BSDTemplateFilePath);
						 break;
					 }
				 }
				 
			 }catch (Exception e)
			 {
				 e.getMessage();
				 System.out.println(e);
			 }
		 }
		 

			
		 
		 public  Boolean  getDataFromDataSheet(String inputData,String SheetName) throws InterruptedException
		 {
			 Boolean excelDataFetched = false;
			
				
			 try 
			 {
				
					String testcaserowNoString = Integer.toString(webActionKeywordsTest.dataSheetTestCaseRowIndex);
				 String dataSheetColName = inputData;
				 String dataSheetCellValue;
				 if(webActionKeywordsTest.inputData.contains("ds="))
				 {
					 
					 dataSheetCellValue= dataSheetColName.split("=")[1].trim();
					 webActionKeywordsTest.dataSheetValue = dataSheetCellValue;
					 if(dataSheetCellValue.equalsIgnoreCase(""))
					 {
						 excelDataFetched = false;
					 }else
					 {
						 excelDataFetched = true;
					 }
					
				 }
				
				 //String dataSheetColName = inputData.split("ds=")[1].trim();
//				 if(dataSheetColName.contains("="))
//				 {
//					 dataSheetColName = dataSheetColName.split("=")[0].trim();
//				 }
//				 if(dataSheetColName.contains(";"))
//				 {
//					 dataSheetColName = dataSheetColName.split(";")[0].trim();
//				 }
				 else {
				  dataSheetCellValue = Get_Excel_ValueHeaderFirst(testcaserowNoString, dataSheetColName, SheetName, Caller.BSDTemplateFilePath);
				 
//				 if(webActionKeywordsTest.inputData.contains("~ds="))
//				 {
//					 webActionKeywordsTest.inputData = webActionKeywordsTest.inputData.split("~")[0] + dataSheetCellValue;
//				 }else if(webActionKeywordsTest.inputData.startsWith("ds="))
//				 if(webActionKeywordsTest.inputData.startsWith("ds="))
//				 {
//					 webActionKeywordsTest.inputData = dataSheetCellValue;
//				 }else if(webActionKeywordsTest.action.equalsIgnoreCase("report"))
//				 {
//					 webActionKeywordsTest.inputData = webActionKeywordsTest.inputData.split("ds=")[0] + "'" + dataSheetCellValue + "'";
//				 }else
//				 {
//					 webActionKeywordsTest.inputData = webActionKeywordsTest.inputData.split("ds=")[0] + dataSheetCellValue;
//				 }
				 
				 webActionKeywordsTest.dataSheetValue = dataSheetCellValue;
				 if(dataSheetCellValue.equalsIgnoreCase(""))
				 {
					 excelDataFetched = false;
				 }else
				 {
					 excelDataFetched = true;
				 }
				 }
				
				 
				 return excelDataFetched;
				 
			 }catch (Exception e)
			 {
				 e.getMessage();
				 System.out.println(e);
			 }
			 
			 return excelDataFetched;
		 }
		 
		//********* Getting the  values from a sheet *****************//
			
			public String Get_Excel_Action_Value(int row_num, int ColValue,String Sheet_Name,String Input_file)
		    {           
		      String value="";
		      
		          try {
		                
		                int Row_Num = 0;              
		                Row_Num = row_num;

		                org.apache.poi.ss.usermodel.Workbook tempWB = null;               
		               // String WB_Name = "C:\\ACES\\ACES_Claimvalidation\\com\\WLP\\AutomationComponents\\Application\\Data\\Input\\TestData.xls";
		                String WB_Name = Input_file;
		                
		                //System.out.println("Excel Filepath:"+ WB_Name);
		                InputStream inp =null;              
		                try {
		                      inp = new FileInputStream(WB_Name);
		                      if(!(WB_Name.contains(".xlsx")))
		                      {
		                    	  tempWB = new HSSFWorkbook(inp);
		                    	  System.out.println(tempWB);
		                      }else
		                      {
		                    	  //tempWB = (org.apache.poi.ss.usermodel.Workbook) new HSSFWorkbook(new POIFSFileSystem(inp));
		                    	  tempWB = new XSSFWorkbook(inp);
		                      }

		                      Sheet s1 = tempWB.getSheet(Sheet_Name);
		                      int Col_Num =ColValue;
		                      Row r = s1.getRow(Row_Num);
		                      
		                      Cell  cel = r.getCell(Col_Num);
		                      
			              		switch (cel.getCellType()) {
			                    case Cell.CELL_TYPE_STRING:
			                    	value = cel.getStringCellValue();
			                   
			                        break;
			                    case Cell.CELL_TYPE_BOOLEAN:
			                    	boolean flag = cel.getBooleanCellValue();
			                    	value=String.valueOf(flag);
			                
			                        break;
			                    case Cell.CELL_TYPE_NUMERIC:
			                    double number = cel.getNumericCellValue();
			                    	value=String.valueOf(number).trim();
			              
			                        break;
			                    case Cell.CELL_TYPE_BLANK:
			                    	value="";
			                    
				                        break;
			              		}
		                    
		                      System.out.println("Value : "+value);
		                      inp.close();
		                      
		                } catch (IOException e) {
		                  System.out.println(e);
		                     
		                      e.getMessage();
		                } catch (Exception e) {
		                      
		                      e.getMessage();
		                     System.out.println(e);
		                }
		                          
		          } catch (Exception e) {
		                System.out.println(" Excel_Read - Function issue : "+e);
		                e.getMessage();
		          }
		            return value;
		                        
		    }
	
			public  Boolean  getDataFromDataSheetUtility(String inputData,String SheetName) throws InterruptedException
			 {
				 Boolean excelDataFetched = false;
				 try 
				 {

						String testcaserowNoString = Integer.toString(1);
					 String dataSheetColName = inputData;
					 //String dataSheetColName = inputData.split("ds=")[1].trim();
//					 if(dataSheetColName.contains("="))
//					 {
//						 dataSheetColName = dataSheetColName.split("=")[0].trim();
//					 }
//					 if(dataSheetColName.contains(";"))
//					 {
//						 dataSheetColName = dataSheetColName.split(";")[0].trim();
//					 }
					 String dataSheetCellValue = Get_Excel_ValueHeaderFirst(testcaserowNoString, dataSheetColName, SheetName, Caller.BSDTemplateFilePath);
					 
//					 if(webActionKeywordsTest.inputData.contains("~ds="))
//					 {
//						 webActionKeywordsTest.inputData = webActionKeywordsTest.inputData.split("~")[0] + dataSheetCellValue;
//					 }else if(webActionKeywordsTest.inputData.startsWith("ds="))
//					 if(webActionKeywordsTest.inputData.startsWith("ds="))
//					 {
//						 webActionKeywordsTest.inputData = dataSheetCellValue;
//					 }else if(webActionKeywordsTest.action.equalsIgnoreCase("report"))
//					 {
//						 webActionKeywordsTest.inputData = webActionKeywordsTest.inputData.split("ds=")[0] + "'" + dataSheetCellValue + "'";
//					 }else
//					 {
//						 webActionKeywordsTest.inputData = webActionKeywordsTest.inputData.split("ds=")[0] + dataSheetCellValue;
//					 }
					 
					 webActionKeywordsTest.dataSheetValue = dataSheetCellValue;
					 if(dataSheetCellValue.equalsIgnoreCase(""))
					 {
						 excelDataFetched = false;
					 }else
					 {
						 excelDataFetched = true;
					 }
					 
					 return excelDataFetched;
					 
				 }catch (Exception e)
				 {
					 e.getMessage();
					 System.out.println(e);
				 }
				 
				 return excelDataFetched;
			 }
			 
			public static int findRowIndexLOB(String row_name, String Test_CaseName,String LOB_Name, int totalRowCount, String Sheet_Name,String Input_file) 
			{
				String value="";
				String Lobvalue="";
				int Row_Num = -1;   
				try 
				{
					org.apache.poi.ss.usermodel.Workbook tempWB = null;               
					// String WB_Name = "C:\\ACES\\ACES_Claimvalidation\\com\\WLP\\AutomationComponents\\Application\\Data\\Input\\TestData.xls";
					String WB_Name = Input_file;

					//System.out.println("Excel Filepath:"+ WB_Name);
					InputStream inp =null;              
					try 
					{
						inp = new FileInputStream(WB_Name);
						if(!(WB_Name.contains(".xlsx")))
						{
							tempWB = new HSSFWorkbook(inp);
							//System.out.println(tempWB);
						}else
						{
							//tempWB = (org.apache.poi.ss.usermodel.Workbook) new HSSFWorkbook(new POIFSFileSystem(inp));
							tempWB = new XSSFWorkbook(inp);
						}

						Sheet s1 = tempWB.getSheet(Sheet_Name);
						int Col_Num =findColNew(s1,Test_CaseName);
                        int Col_lob=findColNew(s1,LOB_Name);
						for(int excelRowNo = 0; excelRowNo<=totalRowCount;excelRowNo++)
						{
							Row r = s1.getRow(excelRowNo);
							Cell  cel = r.getCell(Col_Num);
							Cell lob=r.getCell(Col_lob);
						
							try
							{
								value = cel.toString().trim();
								 Lobvalue=lob.toString().trim();
							} catch(Exception e)
							{
								value = "";
							}
							if(value.equalsIgnoreCase(row_name)&&(Lobvalue.equalsIgnoreCase(Caller.LOB)))
							{
								Row_Num = excelRowNo;
								break;
							}
							
						}
						
						inp.close();
						return Row_Num; 
						
					} catch (IOException e) {
						System.out.println(e);

						e.getMessage();
					} catch (Exception e) {

						e.getMessage();
						System.out.println(e);
					}

				} catch (Exception e) {
					System.out.println(" Excel_Read - Function issue : "+e);
					e.getMessage();
				}
				return Row_Num;  
			}


			public boolean executeNonQueryInAccessDB(String SQLnonQuery)
			{
				boolean SQLnonQueryExecuted = false;
				Connection con=null;
				Statement sta= null;

				try
				{
					Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
					con=DriverManager.getConnection("jdbc:ucanaccess://" + Caller.DataObjectDBpath, "", "");
					sta = con.createStatement();
					sta.execute(SQLnonQuery);
					sta.close();
					con.close();
					SQLnonQueryExecuted = true;

				}catch(Exception e)
				{
					System.out.println("Error in executing executeNonQueryInAccessDB !! ");
					System.out.println(e.toString());
					try
					{
						if (sta != null) sta.close();
						if (con != null) con.close();
					}catch (Exception e1)
					{
						e1.printStackTrace();
						System.out.println(e1.toString());
					}
				}
				return SQLnonQueryExecuted;
			}

			
			public String executeQueryInAccessDBOneRecord(String SQLQuery, String DBcolumnName)
			{
				ResultSet  rs = null;
				Connection con=null;
				Statement sta= null;
				String DBcellData = "";

				try
				{
					Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
					con=DriverManager.getConnection("jdbc:ucanaccess://" + Caller.DataObjectDBpath, "", "");
					sta = con.createStatement();
					rs = sta.executeQuery(SQLQuery);

					while (rs.next()) 
					{
						DBcellData = rs.getString(DBcolumnName).toString();
						break;
					}

					rs.close();
					sta.close();
					con.close();

				}catch(Exception e)
				{
					System.out.println("Error in executing executeQueryInAccessDB !! ");
					System.out.println(e.toString());
					try
					{
						if (rs != null) rs.close();
						if (sta != null) sta.close();
						if (con != null) con.close();
					}catch (Exception e1)
					{
						e1.printStackTrace();
						System.out.println(e1.toString());
					}
				}
				return DBcellData;
			}


		 
		 public  void  getMappedObjPropFromObjSheetByDB(String objectName, String DBcolHeader) throws InterruptedException
			{
				try 
				{
					String SQLQuery = "select * from Objects where ObjectName = '" + objectName + "'";
					String objShtRowNoString = executeQueryInAccessDBOneRecord(SQLQuery, DBcolHeader);
					if(!objShtRowNoString.equalsIgnoreCase(""))
					{
						//int objShtRowNo = Integer.parseInt(objShtRowNoString);
						webActionKeywordsTest.objectName =Get_Excel_ValueHeaderFirst(objShtRowNoString, "ObjectProperty", Caller.objectSheetname, Caller.BSDTemplateFilePath);
					}

				}catch (Exception e)
				{
					e.getMessage();
					System.out.println(e);
					 webActionKeywordsTest.driver.quit();
				}
			}

			

//********* Finding the particular row index in the sheet *****************//
public static int findRowInput(int RowNo,String row_name, int Column_Index, int totalRowCount, String Sheet_Name,String Input_file) 
{
	String value="";
	int Row_Num = -1;   
	try 
	{
		org.apache.poi.ss.usermodel.Workbook tempWB = null;               
		// String WB_Name = "C:\\ACES\\ACES_Claimvalidation\\com\\WLP\\AutomationComponents\\Application\\Data\\Input\\TestData.xls";
		String WB_Name = Input_file;

		//System.out.println("Excel Filepath:"+ WB_Name);
		InputStream inp =null;              
		try 
		{
			inp = new FileInputStream(WB_Name);
			if(!(WB_Name.contains(".xlsx")))
			{
				tempWB = new HSSFWorkbook(inp);
				//System.out.println(tempWB);
			}else
			{
				//tempWB = (org.apache.poi.ss.usermodel.Workbook) new HSSFWorkbook(new POIFSFileSystem(inp));
				tempWB = new XSSFWorkbook(inp);
			}

			Sheet s1 = tempWB.getSheet(Sheet_Name);
			int Col_Num = Column_Index;
			int testcasestartactiontemp=0;
			for(int excelRowNo = RowNo; excelRowNo<=totalRowCount;excelRowNo++)
			{
			
				Row r = s1.getRow(excelRowNo);
				Cell  cel = r.getCell(Col_Num);
				switch (cel.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                	value = cel.getStringCellValue();
                	testcasestartactiontemp++;
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                	boolean flag = cel.getBooleanCellValue();
                	value=String.valueOf(flag);
                	testcasestartactiontemp++;
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                double number = cel.getNumericCellValue();
                	value=String.valueOf(number).trim();
                	testcasestartactiontemp++;
                    break;
                case Cell.CELL_TYPE_BLANK:
                	value="";
                	testcasestartactiontemp++;
                        break;
				}
			
				if(value.equalsIgnoreCase(row_name))
				{
					Row_Num = testcasestartactiontemp;
					break;
				}
				
			}
			
			inp.close();
			return Row_Num; 
			
		} catch (IOException e) {
			System.out.println(e);

			e.getMessage();
		} catch (Exception e) {

			e.getMessage();
			System.out.println(e);
		}

	} catch (Exception e) {
		System.out.println(" Excel_Read - Function issue : "+e);
		e.getMessage();
	}
	return Row_Num;  
}


}


