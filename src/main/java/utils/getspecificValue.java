package utils;

import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class getspecificValue {
	public static void main(String[]args) throws IOException {
		getRowCount();
		getCellData();
	}
	public static void getCellData() throws IOException {
		String excelpath = "./data/shubh.xlsx";
		int i=0,j=0;
		Scanner scanner = new Scanner(System.in);
		System.out.println("enter the row value");
		i = scanner.nextInt();
		System.out.println("enter the column value");
		j = scanner.nextInt();
		
		XSSFWorkbook workbook = new XSSFWorkbook(excelpath);
		XSSFSheet sheet = workbook.getSheet("sheet1");
//		DataFormatter formatter = new DataFormatter();
//		Object value = formatter.formatCellValue(sheet.getRow(1),getCell(2));
		String value=sheet.getRow(i).getCell(j).getStringCellValue();
		System.out.println(value);
	}
	
	public static void getRowCount() {	
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
}
