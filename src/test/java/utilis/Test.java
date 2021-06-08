package utilis;

import java.io.IOException;

public class Test {

	
	public static void main(String[]args) throws IOException {
		String excelPath="./data/TestData.xlsx";
		String sheetName="Sheet1";
		ExcelUtilis excel = new ExcelUtilis(excelPath, sheetName);
		
		excel.getRowCount();
		excel.getCellData();
	}
}
