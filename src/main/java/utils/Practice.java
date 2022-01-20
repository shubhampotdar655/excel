package utils;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class Practice {
	@BeforeTest
	public void test1() {	
		try {
		String excelpath = "./data/shubh.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook(excelpath);
		XSSFSheet sheet = workbook.getSheet("sheet1");
		int rowCount = sheet.getPhysicalNumberOfRows();
		System.out.println("NO of Rows:"+rowCount);
		}catch(Exception exp) {
			System.out.println(exp.getCause());
			System.out.println(exp.getMessage());
			exp.printStackTrace();
		}
	}
	@Test
	public void test2() throws IOException {
		String excelpath = "./data/shubh.xlsx";
		int i=0,j=0;
		Scanner scanner = new Scanner(System.in);
		System.out.println("enter the row value");
		i=1;
		j=1;
//		i = scanner.nextInt();
		System.out.println(i);
		System.out.println("enter the column value");
//		j = scanner.nextInt();
		System.out.println(j);
		 
		XSSFWorkbook workbook = new XSSFWorkbook(excelpath);
		XSSFSheet sheet = workbook.getSheet("sheet1");
//		DataFormatter formatter = new DataFormatter();
//		Object value = formatter.formatCellValue(sheet.getRow(1),getCell(2));
		String value=sheet.getRow(i).getCell(j).getStringCellValue();
		System.out.println(value);
	}
	@AfterTest
	public void test3() {
		System.out.print("Result displayed"); 
	}

}


