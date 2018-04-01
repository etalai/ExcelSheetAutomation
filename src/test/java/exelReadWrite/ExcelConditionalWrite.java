package exelReadWrite;
import java.io.*;
import org.apache.poi.xssf.usermodel.*;
public class ExcelConditionalWrite {
	public static void main(String[] args) throws Exception {
		String excelPath="C:\\Users\\etala\\Desktop\\Book1.xlsx";
//		String excelPath="/src/test/resources/excelsheet/Book1.xlsx";
		FileInputStream in=new FileInputStream(excelPath);
		XSSFWorkbook workbook=new XSSFWorkbook(in);
		XSSFSheet worksheet=workbook.getSheet("Sheet2");
		XSSFCell statusCell=worksheet.getRow(0).getCell(2);
		
		if(statusCell==null) {
			worksheet.getRow(0).createCell(2);
		}
		statusCell.setCellValue("Status");
		
		for(int a=1; a<worksheet.getPhysicalNumberOfRows(); a++) {
			XSSFCell cell=worksheet.getRow(a).createCell(2);
			String value=worksheet.getRow(a).getCell(1).toString().trim();
			if(value.equalsIgnoreCase("python book".trim())) {
				if(cell==null) {
					cell=worksheet.getRow(a).createCell(2);
				}
				cell.setCellValue("Pass");
			}
		}
		
		FileOutputStream out=new FileOutputStream(excelPath);
		workbook.write(out);
		out.close();
		in.close();
	}
}
