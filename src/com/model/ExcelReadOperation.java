package com.model;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadOperation 
{
	public void readExcel(String filename,String sheetname) throws IOException
	{
		//int [][] arrayexcel=null;
		FileInputStream fis=new FileInputStream("E:\\selnium\\Excel_Sheet_Operation\\StudentDetails.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheet(sheetname);
		XSSFRow row=sheet.getRow(2);
		XSSFCell cell=row.getCell(2);
		System.out.println(cell.getStringCellValue());
		
		int rows=sheet.getLastRowNum();
		System.out.println("The no of rows are"+rows);
		
		int rowcount=rows+1;
		System.out.println("The no of rows are"+rowcount);
		
		int columns=sheet.getRow(rows).getLastCellNum();
		System.out.println("The no of columns are:"+columns);
		
		int arrayexcel[][]=new int[rowcount][columns];
		
		for(int i=0;i<rowcount;i++)
		{
			for(int j=0;j<columns;j++)
			{
				System.out.println(sheet.getRow(i).getCell(j));
			}
			
		}
		
		
	}

	public static void main(String[] args) throws IOException 
	{
		ExcelReadOperation ex=new ExcelReadOperation();
		ex.readExcel("E:\\selnium\\Excel_Sheet_Operation\\StudentDetails.xlsx","sheet1");
		

	}

}

/*
 Warje
The no of rows are10
The no of rows are11
The no of columns are:3
Roll_NO
Name_Of_Student
Addresss
1.0
Divya
KarveNagar
2.0
Sonal
Warje
3.0
Rohit
KarveNagar
4.0
Sanika
KarveNagar
5.0
Rupali
Chinchwad
6.0
Akshada
Kothrud
7.0
Saurbh
Kothrud
8.0
Priyanka
Aundh
9.0
Subhodh
KarveNagar
10.0
Nagesh
KarveNagar
*/
