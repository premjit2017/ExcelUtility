package utils;

public class ExcelUtilsTest {

	public static void main(String[] args) {
		
		//String excelPath = "./data/TestData.xlsx";
		String projDir = System.getProperty("user.dir");
		String excelPath = projDir+"/data/TestData.xls";
		String sheetName = "TestData";
		ExcelUtils excel = new ExcelUtils(excelPath, sheetName);
		try {
		excel.getRowCount();
		excel.getCellData(1, 0);
		excel.getCellData(1, 1);
		excel.getCellData(1, 2);
		}catch(Exception e) {}
	}
}
