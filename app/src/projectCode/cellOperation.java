package projectCode;

import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class cellOperation {
	public static void copyCell(XSSFCell srcCell, XSSFCell disCell) throws Exception {
//		if(srcCell == null) {
//			return;
//		}
//		XSSFCellStyle srcCellStyle = srcCell.getCellStyle();
//		XSSFCellStyle disCellStyle = disCell.getRow().getSheet().getWorkbook().createCellStyle();
//		disCellStyle.setBorderBottom(srcCellStyle.getBorderBottom());
//		disCellStyle.setFillForegroundColor(srcCellStyle.getFillForegroundColor());
		disCell.setCellValue(ExcelValue.getValue(srcCell));
//		disCellStyle.cloneStyleFrom(srcCellStyle);
//		disCell.setCellStyle(disCellStyle);
	}
	
	public static void setDateFormat(XSSFCell xssfCell, CellStyle cellStyle) throws Exception{
		SimpleDateFormat simpleDateFormat = new SimpleDateFormat("MM/dd/yyyy");
		if(ExcelValue.getValue(xssfCell).equals("") || ExcelValue.getValue(xssfCell) == null || ExcelValue.getValue(xssfCell).length() != 10) {
			return;
		}
		XSSFWorkbook outputWorkbook = xssfCell.getSheet().getWorkbook();
		if(cellStyle == null) {
			cellStyle = outputWorkbook.createCellStyle();
		}
		XSSFCreationHelper createHelper = outputWorkbook.getCreationHelper();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("MM/dd/yyyy"));
		String cellDateString = ExcelValue.getValue(xssfCell);
		cellDateString = cellDateString.replace(".", "/");
		xssfCell.setCellValue(simpleDateFormat.parse(cellDateString));
		xssfCell.setCellStyle(cellStyle);
	}
	
	public static String getARGBHex(XSSFCell xssfCell) throws Exception{
		if(xssfCell.getCellStyle() == null) {
			return "";
		}
		else if(xssfCell.getCellStyle().getFillForegroundColorColor() == null) {
			return "";
		}
		else {
			return xssfCell.getCellStyle().getFillForegroundColorColor().getARGBHex().substring(2);
		}
	}
}
