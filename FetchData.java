package com.test.ExcelDataFetch;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FetchData
{
	public static void getDataFromExcel() throws IOException

	{
	
	
		FileInputStream fis = new FileInputStream("C:\\Users\\Dilip\\OneDrive\\Documents\\MyRecord.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		Sheet sheet = workbook.getSheet("Dilip");
		System.out.println(sheet.getRow(0).getCell(0).getNumericCellValue());
		System.out.println(sheet.getRow(0).getCell(1).getNumericCellValue());
		System.out.println(sheet.getRow(0).getCell(2).getNumericCellValue());
		System.out.println(sheet.getRow(0).getCell(2).getNumericCellValue());
		fis.close();
	}	
}
