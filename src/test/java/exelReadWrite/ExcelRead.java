package exelReadWrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	public static void main(String[] args) throws Exception {
		String excelPath = "C:\\Users\\etala\\Desktop\\Book1.xlsx";

		// XSSFSheet worksheet = new XSSFWorkbook(new
		// FileInputStream(excelPath)).getSheet("Sheet1");

		FileInputStream in = new FileInputStream(excelPath);
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		XSSFSheet worksheet = workbook.getSheet("Sheet1");

		int rowsCount = worksheet.getPhysicalNumberOfRows();

//		System.out.println(worksheet.getRow(0).getCell(0).toString());
//		System.out.println(worksheet.getRow(1).getCell(0).toString());
//		System.out.println(worksheet.getRow(rowsCount - 1).getCell(1).toString());
		for (int a = 1; a < rowsCount; a++) {

//			System.out.println(
//				worksheet.getRow(a).getCell(0).toString() + "  " 
//				+ worksheet.getRow(a).getCell(1).toString()+ "  " +
//				 worksheet.getRow(a).getCell(2).toString());
			
			System.out.println(worksheet.getRow(a).getCell(1).toString()+" "+
				"works for "+worksheet.getRow(a).getCell(2)+" department");
			
		}
		
		in.close();
	}
}
