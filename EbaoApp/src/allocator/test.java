package allocator;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Set;

import javax.swing.JFileChooser;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;

import businessKeywords.BusinessFunctions;
import businessKeywords.webActionKeywordsTest;



public class test {
	public static webActionKeywordsTest WF = new webActionKeywordsTest();
	public static CommonFunctions CF= new CommonFunctions();
	public static String tableheader="(//select[@class='ms-choice'])[2]";
	public static String tblDisptachAddress="//table[@class='table_head']/tbody/tr[@class='odd']";
	public static String Browser="IEx64";
	
	public static void main(String[] args) throws IOException, InterruptedException {
	    webActionKeywordsTest.driver = CF.launchBrowser(Browser);
		webActionKeywordsTest.driver.get("http://192.168.88.23:7001/ls/login.do");
		webActionKeywordsTest.driver.manage().window().maximize();
		webActionKeywordsTest.objectType="name";
		webActionKeywordsTest.objectProperty="currSettlementOption";
		 WF.getValue();

	}
}