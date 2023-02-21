package com.test.ExcelDataFetch;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SetData {

	public static void SetData() throws IOException
	{
		
	
	{		
		FileInputStream fis = new FileInputStream("C:\\Users\\Dilip\\OneDrive\\Documents\\MyRecord.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet =workbook.getSheet("Dilip");
		XSSFRow row = sheet.createRow(0);
		row.createCell(0).setCellValue(13);
		row.createCell(1).setCellValue(14);
		row.createCell(2).setCellFormula("A1+B1");
		//System.out.println(sheet.getRow(0).getCell(2).getNumericCellValue());
		
		row.createCell(3).setCellFormula("A1+B1+C1");
		FileOutputStream fos = new FileOutputStream("C:\\Users\\Dilip\\OneDrive\\Documents\\MyRecord.xlsx");
		workbook.write(fos);
		fos.close();
		System.out.println("read and write");
	}
	
		
		
	
}

	
}
