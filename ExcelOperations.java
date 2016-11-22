package com.confirmtkt.general.helper;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperations {

	public static Object[][] readExcelData(String path, String sheetName) throws IOException{
		InputStream excelfile = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(excelfile);
		XSSFSheet currentSheet = workbook.getSheet(sheetName);
		int lastRow = currentSheet.getLastRowNum();
		System.out.println("The number of Rows present in sheet is "+lastRow);
		XSSFRow  titleRow = currentSheet.getRow(0);
		int lastCol = titleRow.getLastCellNum();
		System.out.println("The number of columns present in sheet is "+lastCol);
		Object[][] data = new Object[lastRow][lastCol];
		for(int row=1; row<=lastRow; row++){
			XSSFRow currentRow = currentSheet.getRow(row);
			for(int col=0; col<=lastCol-1; col++){
				Cell firstCell = currentRow.getCell(col);
				switch (firstCell.getCellType()) {
				case 0:
					Double d = new Double(firstCell.getNumericCellValue());
					data[row-1][col] = d.intValue();
					break;
				case 1:
					data[row-1][col] = firstCell.getStringCellValue();
					break;
				case 2:                             
					break;                          
				}
			}  
		}
		saveFile(path, workbook);
		return data;
	}
	public static void setData(double data,String path, String sheetName,int colNo,int finalColno) throws IOException{
		
		InputStream excelfile = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(excelfile);
		XSSFSheet currentSheet = workbook.getSheet(sheetName);
		int lastRow = currentSheet.getPhysicalNumberOfRows();
		if (lastRow==1)
			setDateIfLastRow(data, colNo, currentSheet, lastRow);
		else 
			setDateIfNotLastRow(data, colNo, finalColno, currentSheet, lastRow);
		saveFile(path, workbook);
	}
	private static void setDateIfNotLastRow(double data, int colNo, int finalColno, XSSFSheet currentSheet, int lastRow) {
		XSSFRow  row ;
		row = currentSheet.getRow(lastRow-1);
		int lastCol = row.getLastCellNum();
		if (lastCol == finalColno){
			row = currentSheet.createRow(lastRow);
			Cell cell = row.createCell(colNo);
			cell.setCellValue(data);
		}
		else{
			Cell cell = row.createCell(colNo);
			cell.setCellValue(data);
		}
	}
	private static void setDateIfLastRow(double data, int colNo, XSSFSheet currentSheet, int lastRow) {
		System.out.println(lastRow);
		XSSFRow  row ;
		row = currentSheet.createRow(lastRow);
		Cell currCell = row.createCell(colNo);
		currCell.setCellValue(data);
	}
	public static void setData(Date data,String path, String sheetName,int colNo,int finalColno) throws IOException{

		InputStream excelfile = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(excelfile);
		XSSFSheet currentSheet = workbook.getSheet(sheetName);
		int lastRow = currentSheet.getPhysicalNumberOfRows();
		if (lastRow==1)
			setDateIfLastRow(data, colNo, currentSheet, lastRow);
		else 
			setDateIfNotLastRow(data, colNo, finalColno, currentSheet, lastRow);
		saveFile(path, workbook);
	}
	private static void setDateIfNotLastRow(Date data, int colNo, int finalColno, XSSFSheet currentSheet, int lastRow) {
		XSSFRow  row ;
		row = currentSheet.getRow(lastRow-1);
		int lastCol = row.getLastCellNum();
		if (lastCol == finalColno){
			row = currentSheet.createRow(lastRow);
			Cell cell = row.createCell(colNo);
			cell.setCellValue(data);
		}
		else{
			Cell cell = row.createCell(colNo);
			cell.setCellValue(data);
		}
	}
	private static void setDateIfLastRow(Date data, int colNo, XSSFSheet currentSheet, int lastRow) {
		System.out.println(lastRow);
		XSSFRow  row ;
		row = currentSheet.createRow(lastRow);
		Cell currCell = row.createCell(colNo);
		currCell.setCellValue(data);
	}
	public static void setData(String data,String path, String sheetName,int colNo,int finalColno) throws IOException{
		InputStream excelfile = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(excelfile);
		XSSFSheet currentSheet = workbook.getSheet(sheetName);
		int lastRow = currentSheet.getPhysicalNumberOfRows();
		if (lastRow==1)
			setDataIfLastRow(data, colNo, currentSheet, lastRow);
		else 
			setDataIfNotLastRow(data, colNo, finalColno, currentSheet, lastRow);
		saveFile(path, workbook);
	}
	private static void setDataIfNotLastRow(String data, int colNo, int finalColno, XSSFSheet currentSheet,
			int lastRow) {
		XSSFRow  row ;
		row = currentSheet.getRow(lastRow-1);
		int lastCol = row.getLastCellNum();
		System.out.println("Number of Columns " +lastCol);
		if (lastCol == finalColno){
			row = currentSheet.createRow(lastRow);
			Cell cell = row.createCell(colNo);
			cell.setCellValue(data);
		}
		else{
			Cell cell = row.createCell(colNo);
			cell.setCellValue(data);
		}
	}
	private static void setDataIfLastRow(String data, int colNo, XSSFSheet currentSheet, int lastRow) {
		System.out.println(lastRow);
		XSSFRow  row ;
		row = currentSheet.createRow(lastRow);
		Cell currCell = row.createCell(colNo);
		currCell.setCellValue(data);
	}
	public static void saveFile(String filepath,XSSFWorkbook workbook) throws IOException{
		FileOutputStream out = new FileOutputStream( 
		new File(filepath));
		workbook.write(out);
		out.close();
	}
}
