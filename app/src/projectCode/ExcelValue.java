package projectCode;

import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;

public class ExcelValue {
	public static String getValue(XSSFCell xssfcell) {
		if(xssfcell == null) {
			return "";
		}
		else if(xssfcell.getCellType() == CellType.BOOLEAN )
		{
			return String.valueOf(xssfcell.getBooleanCellValue());
		}
		else if(xssfcell.getCellType() == CellType.NUMERIC ) 
		{
			if(HSSFDateUtil.isCellDateFormatted(xssfcell)) {
				SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
				return String.valueOf(sdf.format(xssfcell.getDateCellValue()));
			}
			else {
				xssfcell.setCellType(CellType.STRING);
				return String.valueOf(xssfcell.getStringCellValue());
			}
		}
		else if(xssfcell.getCellType() == CellType.STRING)
		{

			return String.valueOf(xssfcell.getStringCellValue());
		}
		else if(xssfcell.getCellType() == CellType.FORMULA)
		{

			return String.valueOf(xssfcell.getCellFormula());
		}
		else if(xssfcell.getCellType() == CellType.ERROR)
		{

			return String.valueOf(xssfcell.getErrorCellString());
		}
		else
		{
			return null;
			
		}
	}
}
