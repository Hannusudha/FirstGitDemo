package DataExample;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDrive {

	@SuppressWarnings("null")
	public static void main(String[] args) throws IOException {

		
		System.out.println("Updated File");
		FileInputStream fis = new FileInputStream("C:\\Users\\welcome\\Desktop\\DataDriven.xlsx");
		System.out.println(fis);

		XSSFWorkbook wb = new XSSFWorkbook(fis);
		System.out.println(wb);
		XSSFSheet s1 = wb.getSheet("Sheet1");

		int row = s1.getLastRowNum();

		System.out.println("no of rows " + row);

		Row r = s1.getRow(0);
		int col = r.getLastCellNum();
		System.out.println("no of cols " + col);

		Object[][] data = new Object[row][col];
		for (int i = 1; i <= row; row++) {

			for (int j = 0; j < col; j++) {

				data[i - 1][j] = s1.getRow(i).getCell(j).getStringCellValue();
				System.out.print(data[i - 1][j]);
			}
		}

	}

}
