package com.test.ExcelDataFetch;

import java.io.IOException;

import org.testng.annotations.Test;

public class Test1
{
	@Test
	public void fetchData() throws IOException, InterruptedException
	{
//		String data = ExcelSheet.getData("D_Test", 2, 0);
//		System.out.println(data);
		
		//ExcelSheet.updateData("D_Test", 1, 0, "13");
		//ExcelSheet.updateData("D_Test", 1, 1, "14");
		//ReadExcel.getData("D_Test",1,0);
		
		//ExcelSheet.updateData("D_Test", 1, 2, "formula");
		SetData.SetData();
		System.out.println("my number has written");
		Thread.sleep(2000);
		FetchData.getDataFromExcel();
		//System.out.println(ReadExcel.getData("D_Test", 1, 1,"13"));
	}
}
