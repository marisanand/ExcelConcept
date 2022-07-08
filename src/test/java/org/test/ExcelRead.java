package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	public static void main(String[] args) throws Throwable {
		
		File f=new File("C:\\Users\\hp\\eclipse-workspace\\ExcelConcept\\src\\test\\resources\\ExcelWrite.xlsx");
		FileInputStream f2=new FileInputStream(f);
		Workbook w= new XSSFWorkbook(f2);
		Sheet s= w.getSheet("Excel");
		Row r= s.getRow(0);
		Cell c= r.getCell(0);
		int cellType=c.getCellType();
		if(cellType==1) {
			String value=c.getStringCellValue();
			if(value.equals("Mariappan")) {
				c.setCellValue("Anand");
			}
		}
		FileOutputStream f1=new FileOutputStream(f);
		w.write(f1);
	}

}
