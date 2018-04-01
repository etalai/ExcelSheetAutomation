package exelReadWrite;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.*;

public class Amazon {
	private String URL = "jdbc:mysql://localhost:3306/hr";
	private String DbUserName = "root";
	private String DbPassword = "12345";
	private String sql = "SELECT EMPLOYEE_ID, JOB_ID, SALARY FROM EMPLOYEES ORDER BY SALARY DESC;";
	String excelPath = "src/test/resources/excelsheet/Book1.xlsx";

	@Test
	public void setUp() throws Exception {
		Connection connection = DriverManager.getConnection(URL, DbUserName, DbPassword);
		String excelPath = "C:\\Users\\etala\\Desktop\\Book1.xlsx";
		FileInputStream in = new FileInputStream(excelPath);
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		XSSFSheet worksheet = workbook.getSheet("Sheet5");
		
		Statement statement = connection.createStatement();
		ResultSet resultSet = statement.executeQuery(sql);
		ResultSetMetaData rsMetaData=resultSet.getMetaData();
		System.out.println(rsMetaData.getColumnCount());
		System.out.println(rsMetaData.getColumnName(2));
		
		resultSet.last();
		int rowCount = resultSet.getRow();
		resultSet.beforeFirst();
		Map<Integer, String> employee_id = new HashMap<>();
		Map<Integer, String> job_id = new HashMap<>();
		Map<Integer, String> salary = new HashMap<>();

		int num = 1;

		while (resultSet.next()) {
			employee_id.put(num, resultSet.getString("employee_id").toString());
			job_id.put(num, resultSet.getString("job_id").toString());
			salary.put(num, resultSet.getString("salary").toString());
			num++;
		}

		for (int rownum = 0; rownum < rowCount; rownum++) {
			if (rownum == 0) {
				employee_id.put(rownum, "employee_id");
				job_id.put(rownum, "job_id");
				salary.put(rownum, "salary");
			}

			XSSFCell cell = worksheet.getRow(rownum).createCell(0);
			if (cell == null) {
				worksheet.getRow(rownum).createCell(0);
			}
			cell.setCellValue(employee_id.get(rownum));

			XSSFCell cell1 = worksheet.getRow(rownum).createCell(1);
			if (cell1 == null) {
				worksheet.getRow(rownum).createCell(1);
			}
			cell1.setCellValue(job_id.get(rownum));

			XSSFCell cell2 = worksheet.getRow(rownum).createCell(2);
			if (cell2 == null) {
				worksheet.getRow(rownum).createCell(2);
			}
			cell2.setCellValue(salary.get(rownum));

		}

		resultSet.close();
		statement.close();
		connection.close();

		FileOutputStream out = new FileOutputStream(excelPath);
		workbook.write(out);
		out.close();
		in.close();
	}

}
