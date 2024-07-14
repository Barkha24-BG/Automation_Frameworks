package com.testScript.Execution;

import org.apache.poi.xssf.usermodel.XSSFCell;

import com.excel.read.NALExcelXLSReader;
import com.test.actionSteps.CallingActions;

public class Execution {
	int row=0, col=0;
	XSSFCell cell = null;

	public static void startExecution(String sheetName,NALExcelXLSReader reader) {
		String text=reader.getCellData(sheetName, 1, 2);
		System.out.println("text:"+text);
		
	
		int rowCount = reader.getRowCount(sheetName);
		System.out.println("number of rows:"+rowCount);
		int colCount=reader.getColumnCount(sheetName);
		System.out.println("number of columns:"+colCount);
		//System.out.println("empty row in sheet:"+reader.emptyRowCount(sheetName));
		  if(text.equalsIgnoreCase("Start")) { 
			  System.out.println("Start:"+text); 
			  }else{ System.out.println("no execution"+text);
			  }
		 
		System.out.println("steps/actions in sheet:");
		/*
		 * while(i<rowCount) { for(int j=2;j<colCount;j++) {
		 * action=reader.getCellData(sheetName, i, j); // System.out.println(action); }
		 * i++; }
		 */
		int i=1,j=0;
		String action=null,locator=null,data=null,status=null ;
		  while(i<rowCount-1) {
			  if(reader.getCellData(sheetName, i, 2).equalsIgnoreCase("Start")) {
				  System.out.println("***Start of test script***");
			 // if(reader.getCellData(sheetName, i, j)!=null) {  //test script status line checking
				  if(reader.getCellData(sheetName, i, colCount-1).equalsIgnoreCase("No")) {
					  System.out.println("status at"+i+"and"+"j:"+reader.getCellData(sheetName, i, colCount-1));
					  i=reader.nextStart(sheetName, i+1, rowCount-1);
					  System.out.println("***End of test script***");
				  }else {
					  System.out.println("status at of yes or blank - "+i+" and "+"j:"+reader.getCellData(sheetName, i, colCount-1));
					  i++;
				  }
			  }else if(reader.getCellData(sheetName, i, 2).equalsIgnoreCase("End")){
				  System.out.println("***End of test script***");
				  i++;
			  }else {
				  action=reader.getCellData(sheetName, i, 2);
				  locator=reader.getCellData(sheetName, i, 3);
				  data=reader.getCellData(sheetName, i, 4);
				  status=CallingActions.calling(action, locator, data);
				  System.out.print("Not null at row:"+i);
				  System.out.print(" "+action+" "+locator+" "+data+" "+status);
				  
				  i++;
				  }
			  System.out.println();
			  }
			  
			 }
		
}

