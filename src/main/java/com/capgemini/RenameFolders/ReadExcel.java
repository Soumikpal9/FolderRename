package com.capgemini.RenameFolders;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	
	private static final String DIR_NAME = 
			"\\ProjectName\\";
	private static final String APP_NAME_1 = "File.xlsx";
	private static final String APP_NAME_2 = "File.xlsx";
	
	public static void folderRenaming(String dirName, String appName, String ktName) {
		try {
			FileInputStream file = new FileInputStream(new File(dirName + appName));
			Workbook workbook = new XSSFWorkbook(file);
			DataFormatter dataFormatter = new DataFormatter();
			Iterator<Sheet> sheets = workbook.sheetIterator();
			while(sheets.hasNext()) {
				Sheet sh = sheets.next();
				Iterator<Row> rowIterator = sh.iterator();
				
				int rowCount = 0;
				while(rowIterator.hasNext()) {
					Row row = rowIterator.next();
					Iterator<Cell> cellIterator = row.iterator();
					rowCount++;
					if(rowCount > 1) {
						int count = 0;
						String destValue = "", sourceValue = "";
						while(cellIterator.hasNext()) {
							Cell cell = cellIterator.next();
							count++;
							
							if(count == 1) {
								destValue = dataFormatter.formatCellValue(cell);
							}
							File destFile = new File(dirName + destValue);
							
							if(count == 2) {
								sourceValue = dataFormatter.formatCellValue(cell);
							}
							File sourceFile = new File(dirName + sourceValue);
							
							if(!sourceValue.equals("")) {
								if(sourceFile.renameTo(destFile)) {
									System.out.println("Renaming done");
								}
								else {
									System.out.println("Unsuccessful");
								}
							}
							System.out.println("KT dir :     " + dirName+destValue);
							
							String ktFolder = dirName + destValue + "\\" ;
							ktNameRenaming(ktFolder, ktName);
						}
					}
				}
			}
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void ktNameRenaming(String dirName, String ktName) {
		try {
			FileInputStream file = new FileInputStream(new File(dirName + ktName));
			Workbook workbook = new XSSFWorkbook(file);
			DataFormatter dataFormatter = new DataFormatter();
			Iterator<Sheet> sheets = workbook.sheetIterator();
			while(sheets.hasNext()) {
				Sheet sh = sheets.next();
				Iterator<Row> rowIterator = sh.iterator();
				
				int rowCount = 0;
				while(rowIterator.hasNext()) {
					Row row = rowIterator.next();
					Iterator<Cell> cellIterator = row.iterator();
					rowCount++;
					if(rowCount > 1) {
						int count = 0;
						String destValue = "", sourceValue = "";
						while(cellIterator.hasNext()) {
							Cell cell = cellIterator.next();
							count++;
							
							if(count == 2) {
								destValue = dataFormatter.formatCellValue(cell);
							}
							File destFile = new File(dirName + destValue);
							
							if(count == 3) {
								sourceValue = dataFormatter.formatCellValue(cell);
							}
							File sourceFile = new File(dirName + sourceValue);
							
							if(!sourceValue.equals("")) {
								if(sourceFile.renameTo(destFile)) {
									System.out.println("Renaming done");
								}
								else {
									System.out.println("Unsuccessful");
								}
							}
						}
					}
				}
			}
			
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		
		folderRenaming(DIR_NAME, APP_NAME_1, APP_NAME_2);
		
	}
}
