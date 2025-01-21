package tests;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {

		//code to read excel file externally
		//single line comment
		String excelPath = System.getProperty("user.dir") + "\\TutorialsNinjaTestData.xlsx";
		File excelFile = new File(excelPath);
		FileInputStream fis = new FileInputStream(excelFile);

		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("Login");

		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(0).getLastCellNum();

		for (int r = 0; r < rows; r++) {
			XSSFRow row = sheet.getRow(r);

			for (int c = 0; c < cols; c++) {
				XSSFCell cell = row.getCell(c);

				CellType cellType = cell.getCellType();

				switch (cellType) {

				case STRING:
					System.out.print("|" + cell.getStringCellValue() + "|" + " ");
					break;

				case NUMERIC:
					System.out.print((int) cell.getNumericCellValue() + "|" + " ");
					break;

				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue() + " ");
					break;
				}
			}
			System.out.println();
		}
		wb.close();

	}
}
