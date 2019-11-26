package projectCode;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Sort {
	public static int[] insertSort(ArrayList<Integer> processArraylist, int dateIndex, XSSFSheet xssfSheet, boolean monthFirst) throws ParseException {
		int processNum = processArraylist.size();
		int[] sortedProcess = new int[processNum];
		for(int i = 0; i < processNum; i++) {
			sortedProcess[i] = processArraylist.get(i);
		}
		for(int i = 1; i < processNum; i++) {
			Date date = getDate(xssfSheet, dateIndex, sortedProcess[i], monthFirst);
			int current = sortedProcess[i];
			int j = i - 1;
			for(; j>=0 && getDate(xssfSheet, dateIndex, sortedProcess[j], monthFirst).after(date); j--) {
				sortedProcess[j + 1] = sortedProcess[j];
			}
			sortedProcess[j + 1] = current;
		}
		return sortedProcess;
	}
	
	private static Date getDate(XSSFSheet xssfSheet, int dateIndex, int rowIndex, boolean monthFirst) throws ParseException {
		XSSFRow preRow = xssfSheet.getRow(rowIndex);
		XSSFCell dateCell = preRow.getCell(dateIndex);
		String dateString = ExcelValue.getValue(dateCell);
		if(dateString == null || dateString.equals("")) {
			return new Date();
		}
		if(!monthFirst) {
			if(dateString != null && !dateString.equals("")) {
				dateString = dateString.substring(3, 6) + dateString.substring(0, 3) + dateString.substring(6);
			}
		}
		dateString = dateString.replace(".", "/");
		SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
		return sdf.parse(dateString);
	}
}
