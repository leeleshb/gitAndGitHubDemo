package tests;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel1 {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("Sheet1");
		
		Object data[][] = {{"firstName", "lastName", "Age", "City", "Male"},
							{"Leelesh", "Bokde", 34, "Nagpur", true},
							{"Kunal", "Shende", 32, "Mumbai", true},
							{"Bhushan", "Hemane", 34, "Ratnagiri", true},
							{"Sam", "Ingewar", 32, "Mumbai", false}};
		
		int rows = data.length;
		int cols = data[0].length;
		
		for(int r=0; r<rows; r++) {
			XSSFRow row = sheet.createRow(r);
			
			for(int c=0; c<cols; c++) {
				XSSFCell cell = row.createCell(c);
				
				Object cellType = data[r][c];
				
				if(cellType instanceof String) {
					cell.setCellValue((String)cellType);
				}
				else if(cellType instanceof Integer) {
					cell.setCellValue((Integer)cellType);
				}
				else if(cellType instanceof Boolean) {
					cell.setCellValue((Boolean)cellType);
				}
			}
		}
		
		String path = System.getProperty("user.dir")+("\\excel1.xlsx");
		File file = new File(path);
		FileOutputStream fos = new FileOutputStream(file);
		wb.write(fos);
		wb.close();
		System.out.println("Task completes");
	}

}
