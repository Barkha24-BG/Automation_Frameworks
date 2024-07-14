package com.tests;

import com.excel.read.NALExcelXLSReader;
import com.testScript.Execution.Execution;

public class ExcelPOITest {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		NALExcelXLSReader reader = new NALExcelXLSReader("D:\\Automation Script\\Test Script\\Homepage.xlsx");
		String sheetName="Sheet1";
//		int rowCount = reader.getRowCount(sheetName);
//		System.out.println("number of rows:"+rowCount);
//		int colCount=reader.getColumnCount(sheetName);
//		System.out.println("number of columns:"+colCount);
//		if(rowCount==0 && colCount==0) {
//			System.out.println("sheet is empty");
//		}else {
//			System.out.println("sheet has data");
//			String text=reader.getCellData(sheetName,colCount , rowCount);
//			System.out.println("data from excel:"+text);
//		}
		
		Execution.startExecution(sheetName, reader);
	}

}
