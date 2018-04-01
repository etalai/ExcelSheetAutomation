package exelReadWrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelConditionalRead {
	public static void main(String[] args) throws Exception {
		String excelPath="C:\\Users\\etala\\Desktop\\Book1.xlsx";
		FileInputStream in=new FileInputStream(excelPath);
		XSSFWorkbook workbook=new XSSFWorkbook(in);
		XSSFSheet worksheet=workbook.getSheet("Sheet2");
		
		for(int rownum=1; rownum<worksheet.getPhysicalNumberOfRows(); rownum++) {
			String rowCellValue=worksheet.getRow(rownum).getCell(0).toString();			
			if(rowCellValue.equalsIgnoreCase("y")) {
				System.out.println(worksheet.getRow(rownum).getCell(1).toString());
			}		
		}
		
		
		in.close();
	}
}
