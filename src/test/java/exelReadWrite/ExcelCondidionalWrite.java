package exelReadWrite;

import java.io.*;

import org.apache.poi.xssf.usermodel.*;

public class ExcelCondidionalWrite {
				
	public static void main(String[] args) throws Exception{
		String excelPath="C:\\Users\\etala\\Desktop\\Book1.xlsx";
		FileInputStream in=new FileInputStream(excelPath);
		XSSFWorkbook workbook=new XSSFWorkbook(in);
		XSSFSheet worksheet=workbook.getSheet("Sheet3");
		XSSFCell cell=worksheet.createRow(0).createCell(8);
		in.close();
		if(cell==null) {
			worksheet.createRow(0).createCell(8);
		}		
		cell.setCellValue("");
		
		FileOutputStream out=new FileOutputStream(excelPath);
		workbook.write(out);
		out.close();
	}
}
