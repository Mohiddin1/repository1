package com.burts.packg;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel_Read {

	public static void main(String[] args) throws IOException, EncryptedDocumentException, InvalidFormatException {
		 FileInputStream fis = new FileInputStream("G:\\mohiddin\\Subject2\\DemoSource.xlsx");
	        Workbook wb = WorkbookFactory.create(fis);
	        Sheet sh = wb.getSheet("Sheet1");
	        Row row = sh.getRow(1);
	        Cell cell = row.createCell(5);
	        cell.setCellType(Cell.CELL_TYPE_STRING);
	        cell.setCellValue("pass");
	        FileOutputStream fos=new FileOutputStream("G:\\mohiddin\\Subject2\\DemoDestination.xlsx");
	        wb.write(fos);
	}

}
