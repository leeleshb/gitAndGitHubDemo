package tests;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) throws IOException {

		//code to write into excel file
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("Sheet");

		Object data[][] = { { "Name", "City", "Experience" }, 
				{ "Leelesh", "Nagpur", 3 }, 
				{ "Minu", "Nagpur", 2 },
				{ "Samu", "Pune", 3 } };

		int rows = data.length;
		int cols = data[0].length;

		for (int r = 0; r < rows; r++) {
			XSSFRow row = sheet.createRow(r);
			for (int c = 0; c < cols; c++) {
				XSSFCell cell = row.createCell(c);

				Object cellValue = data[r][c];

				if (cellValue instanceof String) {
					cell.setCellValue((String) cellValue);
				} else if (cellValue instanceof Integer) {
					cell.setCellValue((int) cellValue);
				} else if (cellValue instanceof Boolean) {
					cell.setCellValue((Boolean) cellValue);
				}
			}
		}
		String path = System.getProperty("user.dir") + ("//excel.xlsx");
		File file = new File(path);
		FileOutputStream fos = new FileOutputStream(file);
		wb.write(fos);
		wb.close();
		System.out.println("task completes");

	}

}
