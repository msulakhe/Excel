package utilis;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtilis {
	
	public ExcelUtilis(String excelPath, String sheetName) {
		// TODO Auto-generated constructor stub
	}

	public static void main(String[] args) throws IOException {
		getRowCount();
		getCellData();
		
	}
	
	public static void getCellData() throws IOException {
		
		String excelPath = "./data/Test.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
		
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		String value = sheet.getRow(0).getCell(1).getStringCellValue();
		System.out.println(value);
		
	}

	public static void getRowCount() {
		try {
		String excelPath = "./data/Test.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
		
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		int rowCount = sheet.getPhysicalNumberOfRows();
		System.out.println("No.of Rows"+rowCount);
		}catch(Exception exp) {
			System.out.println(exp.getCause());
			System.out.println(exp.getMessage());
			exp.printStackTrace();
			
		}
	}
	
	
	
}
