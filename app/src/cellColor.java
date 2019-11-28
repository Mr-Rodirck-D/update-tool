import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import projectCode.cellOperation;

public class cellColor {
	public static void main(String[] args) {
		String filePath = "/Users/rodrick.d/Downloads/Greens.xlsx";
		File colorFile = new File(filePath);
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(colorFile);
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		XSSFSheet sheet = workbook.getSheetAt(0);
		int rowNum = sheet.getLastRowNum() + 1;
		for(int i = 0; i < rowNum; i++) {
			XSSFRow row = sheet.getRow(i);
			XSSFCell cell = row.getCell(1);
			if(cell != null) {
				try {
					System.out.println(cellOperation.getARGBHex(cell));
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		try {
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
