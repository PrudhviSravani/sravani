package org.cts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BaseClass {
	public static String getDataFromExcel(int rowNo, int cellNo) throws IOException {
		String value =null;
		File loc = new File("C:\\Users\\Dell\\Desktop\\Prudhvi\\ExcelIntegration\\excel\\TestData 1.xlsx");
		FileInputStream stream = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("srav");
		Row r = s.getRow(rowNo);
		Cell c= r.getCell(cellNo);
		int type = c.getCellType();
		if(type==1) {
			value = c.getStringCellValue();
		}
		else if (type==0) {
			if(DateUtil.isCellDateFormatted(c));
			Date dateCellValue = c.getDateCellValue();
			SimpleDateFormat sim = new SimpleDateFormat("dd-MMM-yy");
			value = sim.format(dateCellValue);
		}
		else {
			double numericCellValue = c.getNumericCellValue();
			long l = (long)numericCellValue;
			value = String.valueOf(numericCellValue);
			
			
			
		}
		
		
	

	
	


		return value;
	}

}
