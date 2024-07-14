package com.excel.read;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Iterator;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class NALExcelXLSReader {
	public String path;
	public FileInputStream fis = null;
	public FileOutputStream fileOut = null;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private XSSFRow row = null;
	private XSSFCell cell = null;

	public NALExcelXLSReader(String path) {

		this.path = path;
		try {

			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			fis.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//number of rows in sheet
	public int getRowCount(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return 0;
		else {
			System.out.println("Index of sheet:"+index);
			sheet = workbook.getSheetAt(index);
			/*
			 * int number = sheet.getLastRowNum() + 1; return number;
			 */
			int count=0;
			//row = sheet.getRow(0);
			while(count>=0) {
				if (sheet.getRow(count) != null) {
					count++;
				}else {
					break;
				}
			}
			return count+1;
		}

	}
	
	//number of column in sheet
	public int getColumnCount(String sheetName) {
		// check if sheet exists
		if (!isSheetExist(sheetName))
			return -1;

		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(0);

		if (row == null)
			return 0;

		return row.getLastCellNum();

	}
	
	//checking sheet exists or not
	public boolean isSheetExist(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1) {
			index = workbook.getSheetIndex(sheetName.toUpperCase());
			if (index == -1)
				return false;
			else
				return true;
		} else
			return true;
	}
	
	/**
	 * returns the data from a cell
	 * @param sheetName
	 * @param colNum
	 * @param rowNum
	 * @return
	 */
	public String getCellData(String sheetName, int rowNum, int colNum) {
		try {
			int index = workbook.getSheetIndex(sheetName);
			String valString=String.valueOf(index);
			if (index == -1)
				return "No sheet found";
			sheet = workbook.getSheet(sheetName);
			row = sheet.getRow(rowNum);
			cell = row.getCell(colNum);
			if (cell==null) {
				return null;
			}else {
			String data=cell.getStringCellValue();
			return data;
			}
			
		}catch(Exception e) {
			return e.getMessage();
		}		
	}
	public int nextStart(String sheetName, int start,int rowNum) {
		try {	
			
				while(start<=rowNum) {
					if(getCellData(sheetName, start, 0)==null) 
						start++;
					else
						break;
				}
			System.out.println("Next start should be:"+start);
			return start;
		}catch(Exception e) {
			return 0;
		}		
	}

	
}
