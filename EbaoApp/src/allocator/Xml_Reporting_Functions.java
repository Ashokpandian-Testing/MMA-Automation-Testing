package allocator;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.w3c.dom.Attr;
import org.w3c.dom.Element;

import businessKeywords.webActionKeywordsTest;

public class Xml_Reporting_Functions {
	
	 


	public void CreateCustomXml(){
			try {
				
				 Calendar calendar = Calendar.getInstance();
				    Date POSStartTime = calendar.getTime();				    
				    SimpleDateFormat POSStartTimeFormat = new SimpleDateFormat("ddMMyyyy_HHmmss");
		       	    String Creationtime = POSStartTimeFormat.format(POSStartTime);
		       	    
		       	 SimpleDateFormat dateFormat2 = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");		     	
		       	 	POSStartTime = calendar.getTime();
		         String strstartTime = dateFormat2.format(POSStartTime); 
		         String Tempstring = strstartTime.replace(":","");
		     	String TempDate[]=Tempstring.split(" ",2);
		     	String BSDSummaryReportPath=Caller.BSDExecutionResultPath+"\\"+TempDate[0];
		     	
		         File oldFile= new File(BSDSummaryReportPath);
		     	 if((!oldFile.isDirectory())) {
		     		 oldFile.mkdir();
		     	    }
		       	    
		       	 BSDSummaryReportPath=Caller.BSDExecutionResultPath+"\\"+TempDate[0];
				String strPassFail = "P";
				Caller.xmlFilePath = BSDSummaryReportPath+"\\"+Caller.BSDNo+"\\"+strPassFail+"_"+Caller.BSDNo+"_"+Creationtime+".xml";

	            DocumentBuilderFactory documentFactory = DocumentBuilderFactory.newInstance();

	            DocumentBuilder documentBuilder = documentFactory.newDocumentBuilder();

	            Caller.document = documentBuilder.newDocument();
	            
//	            logfile.writeLine("<?xml-stylesheet href='Report.xsl' type='text/xsl'?>")
//	           document.createProcessingInstruction("xml-stylesheet", "type=\"text/xsl\" href=\"Report.xsl\"");
//	            Document document = addingStylesheet(doc1);
	            // root element
	            
	            Caller.document.setXmlStandalone(true);
	            org.w3c.dom.ProcessingInstruction pi = Caller.document.createProcessingInstruction("xml-stylesheet", "type=\"text/xsl\" href=\"Report.xsl\"");

	            Element root = Caller.document.createElement("Report");
	            Caller.document.appendChild(root);
	            Caller.document.insertBefore((org.w3c.dom.Node) pi, root); 
	  
//	            Element root = document.createElement("Report");

//	            document.appendChild(root);

	 

	            // employee element

	            Caller.BSDNode = Caller.document.createElement("TestSuite");

	            Caller.BSDNode.appendChild(Caller.document.createTextNode(Caller.BSDNo));

	            root.appendChild(Caller.BSDNode);

	 

	            // set an attribute to staff element
	            Attr attr = Caller.document.createAttribute("StartTime");
	            
	            SimpleDateFormat POSCreationTimeFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");
	       	    Creationtime = POSCreationTimeFormat.format(POSStartTime);

	            attr.setValue(Creationtime);

	            Caller.BSDNode.setAttributeNode(attr);

	            Attr attr1 = Caller.document.createAttribute("Desc");

	            attr1.setValue(Caller.BSDNo);

	            Caller.BSDNode.setAttributeNode(attr1);
	            
	            Attr attr2 = Caller.document.createAttribute("EndTime");

	            attr2.setValue(Creationtime);

	            Caller.BSDNode.setAttributeNode(attr2);
	            	     	 

	            // create the xml file

	            //transform the DOM Object to an XML File

	            TransformerFactory transformerFactory = TransformerFactory.newInstance();

	            Transformer transformer = transformerFactory.newTransformer();

	            DOMSource domSource = new DOMSource(Caller.document);

	            StreamResult streamResult = new StreamResult(new File(Caller.xmlFilePath).getPath());

	 

	            // If you use

//	            StreamResult result = new StreamResult(System.out);

	            // the output will be pushed to the standard output ...

	            // You can use that for debugging

	            transformer.transform(domSource, streamResult);

	            System.out.println("Done creating XML File");

	 

	        } catch (ParserConfigurationException pce) {

	            pce.printStackTrace();

	        } catch (TransformerException tfe) {

	            tfe.printStackTrace();

	        }
	 }
	
	
	
	public void CreateCustomXmlNew(){
		try {
			
			 Calendar calendar = Calendar.getInstance();
			    Date POSStartTime = calendar.getTime();				    
			    SimpleDateFormat POSStartTimeFormat = new SimpleDateFormat("dd_MM_yyyy-HHmmss");
	       	    String Creationtime = POSStartTimeFormat.format(POSStartTime);
	       	    
	       	 SimpleDateFormat dateFormat2 = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");		     	
	       	 	POSStartTime = calendar.getTime();
	         String strstartTime = dateFormat2.format(POSStartTime); 
	         String Tempstring = strstartTime.replace(":","");
	     	String TempDate[]=Tempstring.split(" ",2);
	     	String BSDSummaryReportPath="C:\\Automation Results\\";
	     	
	         File oldFile= new File(BSDSummaryReportPath);
	     	 if((!oldFile.isDirectory())) {
	     		 oldFile.mkdir();
	     	    }
	       	    
//	       	 BSDSummaryReportPath=Caller.BSDExecutionResultPath+"\\"+TempDate[0];
			//String strPassFail = "P";
			Caller.xmlFilePath = BSDSummaryReportPath+Caller.BSDNo+"_"+Creationtime+".xml";

            DocumentBuilderFactory documentFactory = DocumentBuilderFactory.newInstance();

            DocumentBuilder documentBuilder = documentFactory.newDocumentBuilder();

            Caller.document = documentBuilder.newDocument();
            
//            logfile.writeLine("<?xml-stylesheet href='Report.xsl' type='text/xsl'?>")
//           document.createProcessingInstruction("xml-stylesheet", "type=\"text/xsl\" href=\"Report.xsl\"");
//            Document document = addingStylesheet(doc1);
            // root element
            
            Caller.document.setXmlStandalone(true);
            org.w3c.dom.ProcessingInstruction pi = Caller.document.createProcessingInstruction("xml-stylesheet", "type=\"text/xsl\" href=\"Report.xsl\"");

            Element root = Caller.document.createElement("Report");
            Caller.document.appendChild(root);
            Caller.document.insertBefore((org.w3c.dom.Node) pi, root); 
  
//            Element root = document.createElement("Report");

//            document.appendChild(root);

 

            // employee element

            Caller.BSDNode = Caller.document.createElement("TestSuite");

            Caller.BSDNode.appendChild(Caller.document.createTextNode(Caller.BSDNo));

            root.appendChild(Caller.BSDNode);

 

            // set an attribute to staff element
            Attr attr = Caller.document.createAttribute("StartTime");
            
            SimpleDateFormat POSCreationTimeFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");
       	    Creationtime = POSCreationTimeFormat.format(POSStartTime);

            attr.setValue(Creationtime);

            Caller.BSDNode.setAttributeNode(attr);

            Attr attr1 = Caller.document.createAttribute("Desc");

            attr1.setValue(Caller.BSDNo);

            Caller.BSDNode.setAttributeNode(attr1);
            
//            Attr attr2 = Caller.document.createAttribute("EndTime");
//
//            attr2.setValue(Creationtime);
//
//            Caller.BSDNode.setAttributeNode(attr2);
            	     	 

            // create the xml file

            //transform the DOM Object to an XML File

            TransformerFactory transformerFactory = TransformerFactory.newInstance();

            Transformer transformer = transformerFactory.newTransformer();

            DOMSource domSource = new DOMSource(Caller.document);

            StreamResult streamResult = new StreamResult(new File(Caller.xmlFilePath).getPath());

 

            // If you use

//            StreamResult result = new StreamResult(System.out);

            // the output will be pushed to the standard output ...

            // You can use that for debugging

            transformer.transform(domSource, streamResult);

            System.out.println("Done creating XML File");

 

        } catch (ParserConfigurationException pce) {

            pce.printStackTrace();

        } catch (TransformerException tfe) {

            tfe.printStackTrace();

        }
 }
	 
	 public void AddLOBEndTime(){
		 try
		 {
			 Calendar calendar = Calendar.getInstance();
			    Date POSEndTime = calendar.getTime();				    			  
			 SimpleDateFormat POSCreationTimeFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");
	       	    String LobEndtime = POSCreationTimeFormat.format(POSEndTime);
	       	 Attr attr2 = Caller.document.createAttribute("EndTime");

	         attr2.setValue(LobEndtime);

	         Caller.BSDNode.setAttributeNode(attr2);
	         
	         TransformerFactory transformerFactory = TransformerFactory.newInstance();

	            Transformer transformer = transformerFactory.newTransformer();

	            DOMSource domSource = new DOMSource(Caller.document);

	            StreamResult streamResult = new StreamResult(new File(Caller.xmlFilePath).getPath());

	            transformer.transform(domSource, streamResult);

	            System.out.println("End Time Updated");
		 }  catch (TransformerException tfe) {

	            tfe.printStackTrace();

	        }
	
    
		 }
    
	 public void AddLobSummaryXml(){
		 try
		 {
		 Element StepResult = Caller.document.createElement("LOBRow");

		 StepResult.appendChild(Caller.document.createTextNode("LOBSummary"));

		 Caller.BSDNode.appendChild(StepResult);

         // set an attribute to staff element
         Attr attr4 = Caller.document.createAttribute("LOBName");

         attr4.setValue(Caller.LOB);

         StepResult.setAttributeNode(attr4);
         
         Attr attr5 = Caller.document.createAttribute("TotalTest");

         attr5.setValue(Integer.toString(Caller.TotalTestCaseCount));

         StepResult.setAttributeNode(attr5);
         
         Attr attr6 = Caller.document.createAttribute("Executed");

         attr6.setValue(Integer.toString(Caller.TotalExecuted));

         StepResult.setAttributeNode(attr6);
         
         Attr attr7 = Caller.document.createAttribute("Passed");

         attr7.setValue(Integer.toString(Caller.TotalPassed));

         StepResult.setAttributeNode(attr7);
         
         Attr attr8 = Caller.document.createAttribute("Failed");

         attr8.setValue(Integer.toString(Caller.TotalFailed));

         StepResult.setAttributeNode(attr8);
         
         Attr attr9 = Caller.document.createAttribute("NotExecuted");

         attr9.setValue(Integer.toString(Caller.TotalNotExecuted));

         StepResult.setAttributeNode(attr9);
         
         Attr attr10 = Caller.document.createAttribute("PassPer");

         attr10.setValue(Integer.toString(Caller.ExecutionPercentage)+" %");

         StepResult.setAttributeNode(attr10);
         
         Attr attr11 = Caller.document.createAttribute("FailPer");

         attr11.setValue(Integer.toString(Caller.FailedPercentage)+" %");

         StepResult.setAttributeNode(attr11);                
         
         TransformerFactory transformerFactory = TransformerFactory.newInstance();

         Transformer transformer = transformerFactory.newTransformer();

         DOMSource domSource = new DOMSource(Caller.document);

         StreamResult streamResult = new StreamResult(new File(Caller.xmlFilePath).getPath());

         transformer.transform(domSource, streamResult);

         System.out.println("XML LOB Dash Updated");

     } catch (TransformerException tfe) {

         tfe.printStackTrace();

     }
	 }
	 
	 
	 
	 
	 
	 public void AddResult(String Result,String Status){
		 
		 try 
		 {
		 Caller.TCStepResult = Caller.document.createElement("Result");

		 Caller.TCStepResult.appendChild(Caller.document.createTextNode(Result));

		 Caller.TestCase.appendChild(Caller.TCStepResult);



         // set an attribute to staff element
         Attr attr4 = Caller.document.createAttribute("Status");

         attr4.setValue(Status);

         Caller.TCStepResult.setAttributeNode(attr4);
                
         
         TransformerFactory transformerFactory = TransformerFactory.newInstance();

         Transformer transformer = transformerFactory.newTransformer();

         DOMSource domSource = new DOMSource(Caller.document);

         StreamResult streamResult = new StreamResult(new File(Caller.xmlFilePath).getPath());

         transformer.transform(domSource, streamResult);

         System.out.println("XML Updated");
         
         
         
         if(Caller.BRRuleReference != null && !Caller.BRRuleReference.isEmpty())
         {
        	 
        	 Element BRResult = Caller.document.createElement("BR");
        	 
        	 if (Status == "1") {
        		 BRResult.appendChild(Caller.document.createTextNode("Business Rule : "+Caller.BRRuleReference+" has been Passed"));
        	 }
        	 else
        	 {
        		 BRResult.appendChild(Caller.document.createTextNode("Business Rule : "+Caller.BRRuleReference+" has been Failed"));
        	 } 

        	 Caller.TestCase.appendChild(BRResult);


        	 // set an attribute to BR Element
             Attr attr5 = Caller.document.createAttribute("Status");

             attr5.setValue(Status);

             BRResult.setAttributeNode(attr5);
         }
        
         

         transformer.transform(domSource, streamResult);

         System.out.println("BR Updated");



     } catch (TransformerException tfe) {

         tfe.printStackTrace();

     }

         
	 }
	 
public static void AddScreenshot(boolean Status,String Screenshotpath){
		 
Calendar calendar = Calendar.getInstance();
 Date POSStartTime = calendar.getTime();				    
 SimpleDateFormat POSStartTimeFormat = new SimpleDateFormat("ddMMyyyy_HHmmss");
 String Creationtime = POSStartTimeFormat.format(POSStartTime);
		 try 
		 {
		 Element StepResult = Caller.document.createElement("Result");

		 StepResult.appendChild(Caller.document.createTextNode(""));

		 Caller.TestCase.appendChild(StepResult);
		 
		 
//		 String PassedScreenPath = +webActionKeywordsTest.PageName+"_"+webActionKeywordsTest.ReportObjectName+fileIndex+TempDate[0]+".jpg"
//    	 File p= new File(PassedScreenPath);
//    	String FailedScreenPath = Caller.InputDir+"\\Results\\_errorImages\\";
//    	 File f= new File(FailedScreenPath);
if 	(Status){

         // set an attribute to staff element
         Attr attr4 = Caller.document.createAttribute("ScreenShotPath");
        
    	 
    	    
//    	 try {
//
//    	  if((!p.isDirectory())) {
//		    	p.getParentFile().mkdirs();
//		    }
//
//    	}catch(Exception e) {
//    		e.getMessage();
//         System.out.println(e);
//    		
//    	}
		  
         attr4.setValue(Screenshotpath);

         StepResult.setAttributeNode(attr4);
}
else
{
// set an attribute to staff element
Attr attr4 = Caller.document.createAttribute("ErrorScreenShotPath");

// if((!f.isDirectory())) {
// 	f.getParentFile().mkdirs();
// }
attr4.setValue(Screenshotpath);

StepResult.setAttributeNode(attr4);

}

         
         
         
         TransformerFactory transformerFactory = TransformerFactory.newInstance();

         Transformer transformer = transformerFactory.newTransformer();

         DOMSource domSource = new DOMSource(Caller.document);

         StreamResult streamResult = new StreamResult(new File(Caller.xmlFilePath).getPath());

         transformer.transform(domSource, streamResult);

         System.out.println("XML Updated");



     } catch (TransformerException tfe) {

         tfe.printStackTrace();

     }

         
	 }
	 
	 
public void AddLOBNode(){
		 
		 try 
		 {
			 Caller.LOBNode = Caller.document.createElement("TestCase");

	            Caller.LOBNode.appendChild(Caller.document.createTextNode(Caller.LOB));

	            Caller.BSDNode.appendChild(Caller.LOBNode);

	 

	            // set an attribute to staff element
	            Attr attr8 = Caller.document.createAttribute("StartTime");
	            Calendar calendar = Calendar.getInstance();
			    Date LOBStartTime = calendar.getTime();				    
			    SimpleDateFormat XmlCreationtime = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");
	       	    String Creationtime = XmlCreationtime.format(LOBStartTime);		       	    
	            attr8.setValue(Creationtime);

	            Caller.LOBNode.setAttributeNode(attr8);

	            Attr attr9 = Caller.document.createAttribute("Desc");

	            attr9.setValue(Caller.LOB);

	            Caller.LOBNode.setAttributeNode(attr9);
	            
	            Attr attr10 = Caller.document.createAttribute("Status");

	            attr10.setValue("1");

	            Caller.LOBNode.setAttributeNode(attr10);
	            
	            Attr attr11 = Caller.document.createAttribute("EndTime");

	            attr11.setValue(Creationtime);

	            Caller.LOBNode.setAttributeNode(attr11);

	            TransformerFactory transformerFactory = TransformerFactory.newInstance();

	            Transformer transformer = transformerFactory.newTransformer();

	            DOMSource domSource = new DOMSource(Caller.document);

	            StreamResult streamResult = new StreamResult(new File(Caller.xmlFilePath).getPath());

	            transformer.transform(domSource, streamResult);

	            System.out.println("LOB Node Added");

     } catch (TransformerException tfe) {

         tfe.printStackTrace();

     }

         
	 }	 
	 
	 
public void AddTestCaseNode(){
		 
		 try 
		 {
			 Caller.TestCase = Caller.document.createElement("BP");

	            Caller.TestCase.appendChild(Caller.document.createTextNode(webActionKeywordsTest.testCaseID));

	            Caller.LOBNode.appendChild(Caller.TestCase);

	 

	            // set an attribute to staff element
	            Attr attr4 = Caller.document.createAttribute("StartTime");
	            Calendar calendar = Calendar.getInstance();
			    Date TestCaseStartTime = calendar.getTime();				    
			    SimpleDateFormat XmlCreationtime = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");
	       	    String Creationtime = XmlCreationtime.format(TestCaseStartTime);		       	    
	            attr4.setValue(Creationtime);

	            Caller.TestCase.setAttributeNode(attr4);

	            Attr attr5 = Caller.document.createAttribute("Desc");

	            attr5.setValue(webActionKeywordsTest.testCaseID);

	            Caller.TestCase.setAttributeNode(attr5);
	            
//	            Attr attr6 = Caller.document.createAttribute("Status");
//
//	            attr6.setValue("1");
//
//	            TestCase.setAttributeNode(attr6);
	            
	            Attr attr7 = Caller.document.createAttribute("EndTime");

	            attr7.setValue(Creationtime);

	            Caller.TestCase.setAttributeNode(attr7);

	            TransformerFactory transformerFactory = TransformerFactory.newInstance();

	            Transformer transformer = transformerFactory.newTransformer();

	            DOMSource domSource = new DOMSource(Caller.document);

	            StreamResult streamResult = new StreamResult(new File(Caller.xmlFilePath).getPath());

	            transformer.transform(domSource, streamResult);

	            System.out.println("Test Case Node Added");

     } catch (TransformerException tfe) {

         tfe.printStackTrace();

     }

         
	 }


public void AddTestCaseStatusAtt(){
	 
	 try 
	 {
		  Attr attr5 = Caller.document.createAttribute("Status");
		 String Execution = Caller.StringExecutionStatus.toString();
         
         if(Execution.contains("FAIL")) {
        	 attr5.setValue("3");
         }
         else {
             attr5.setValue("1");
         }


           Caller.TestCase.setAttributeNode(attr5);
           

           TransformerFactory transformerFactory = TransformerFactory.newInstance();

           Transformer transformer = transformerFactory.newTransformer();

           DOMSource domSource = new DOMSource(Caller.document);

           StreamResult streamResult = new StreamResult(new File(Caller.xmlFilePath).getPath());

           transformer.transform(domSource, streamResult);

           System.out.println("Test Case Status attrib Node");

} catch (TransformerException tfe) {

    tfe.printStackTrace();

}

    
}


public static void XmlscreenShot(boolean Status) 
{ 
		

	String ScreenshotPath="C:\\Automation Results\\Screenshots\\";
	try {

	    
		//BSDFinalResultPath ="D:\\MMAProject\\MMA_Automation\\Results"+"\\"+TempDate[0]+"\\"+"BSD_013"+"\\"+"BSD_013"+"_Result_"+Tempstring+".xlsx";
	    File f= new File(ScreenshotPath);
	    if((!f.isDirectory())) {
	    	f.mkdirs();
	    }


	
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

    int fileIndex = 1;
    
	File XmlScreenshot = new File(ScreenshotPath+Caller.BSDNo+"_"+Caller.LOB+"_"+webActionKeywordsTest.testCaseID+"_"+webActionKeywordsTest.PageName+"_"+webActionKeywordsTest.ReportObjectName+fileIndex+TempDate[0]+".png");
    String StringScreenshot = ScreenshotPath+Caller.BSDNo+"_"+Caller.LOB+"_"+webActionKeywordsTest.testCaseID+"_"+webActionKeywordsTest.PageName+"_"+webActionKeywordsTest.ReportObjectName+fileIndex+TempDate[0]+".png";
    if(webActionKeywordsTest.driver instanceof InternetExplorerDriver){
        if(isScrollBarPresent){
            while(scrollHeight > 0){

            File srcFile = ((TakesScreenshot)webActionKeywordsTest.driver).getScreenshotAs(OutputType.FILE);
            
                org.apache.commons.io.FileUtils.copyFile(srcFile, XmlScreenshot);
                AddScreenshot(Status,StringScreenshot);
                jexec.executeScript("window.scrollTo(0,"+clientHeight*fileIndex++ +")");
                scrollHeight = scrollHeight - clientHeight;
               
                                  
                                       
        }
        }else{
            File srcFile = ((TakesScreenshot)webActionKeywordsTest.driver).getScreenshotAs(OutputType.FILE);
            org.apache.commons.io.FileUtils.copyFile(srcFile,XmlScreenshot);
            AddScreenshot(Status,StringScreenshot);
        }
    }else{
        File srcFile = ((TakesScreenshot)webActionKeywordsTest.driver).getScreenshotAs(OutputType.FILE);
        org.apache.commons.io.FileUtils.copyFile(srcFile,XmlScreenshot);
        AddScreenshot(Status,StringScreenshot);
    }
    
	}catch(IOException e) {
		
	}

}


}
