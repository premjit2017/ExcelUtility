package utils;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	
//	static XSSFWorkbook workbook;
//	static XSSFSheet sheet;
	static HSSFWorkbook workbook;
	static HSSFSheet sheet;
	
	public ExcelUtils(String excelPath, String sheetName) {
		try {
			
			InputStream file = new FileInputStream(excelPath);
			workbook = new HSSFWorkbook(new POIFSFileSystem(file));
				sheet = workbook.getSheet(sheetName);
		}catch(Exception exp) {
			System.out.println(exp.getCause());
			System.out.println(exp.getMessage());
			exp.printStackTrace();
		}
	}
/*	public static void main(String[] args) throws Exception {
		getRowCount();
		getCellData();
	}*/
	
	public static void getCellData(int rowNum, int colNum) throws IOException {
	//	try {
//			String excelPath = "./data/TestData.xlsx";
//			XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
//			XSSFSheet sheet = workbook.getSheet("TestData");
			
			DataFormatter formatter = new DataFormatter();
			Object value = formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
			//String value = sheet.getRow(1).getCell(2).getStringCellValue();
			System.out.println(value);
		//}catch(Exception exp) {
//			System.out.println(exp.getCause());
//			System.out.println(exp.getMessage());
//			exp.printStackTrace();
//		}
		
	}
	
	public static void getRowCount() {
	//	try {
//		String projDir = System.getProperty("user.dir");
//		System.out.println(projDir);
//		
//		String excelPath = "./data/TestData.xlsx";
//		XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
//		XSSFSheet sheet = workbook.getSheet("TestData");
		int rowCount = sheet.getPhysicalNumberOfRows();
		System.out.println("No of Rows: " +rowCount);
//	}catch(Exception exp) {
//		System.out.println(exp.getCause());
//		System.out.println(exp.getMessage());
//		exp.printStackTrace();
//	}
	}

}
