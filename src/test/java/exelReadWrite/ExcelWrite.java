package exelReadWrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
	public static void main(String[] args) throws Exception {
		String excelPath = "C:\\Users\\etala\\Desktop\\Book1.xlsx";
		FileInputStream in = new FileInputStream(excelPath);
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		XSSFSheet worksheet = workbook.getSheet("Sheet2");
		XSSFCell cell = worksheet.getRow(1).getCell(3);
		
		for(int a=1; a<worksheet.getPhysicalNumberOfRows(); a++) {
			String items=worksheet.getRow(a).getCell(1).toString();
			cell=worksheet.getRow(a).getCell(2);
			if(items.equalsIgnoreCase("Android")) {
				if(cell==null) {
					cell=worksheet.getRow(a).createCell(2);
				}
				cell.setCellValue("azyr");
			}
		}
		
		if (cell == null) {
			cell = worksheet.getRow(1).createCell(2);
		}
		cell.setCellValue("Pass");

		// belows are written at the bottom after everything
		FileOutputStream out = new FileOutputStream(excelPath);
		workbook.write(out);
		in.close();
		out.close();

	}
}
