package projectCode;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class updateEngine {
	public static int addRows = 0;
	public static int deleteRows = 0;
	private static boolean testFlag = false;
	public static String updateEngine(String originalPath, String formatPath) throws Exception{
		//change the file path format
		originalPath = originalPath.replace("\\", "/");
		formatPath = formatPath.replace("\\", "/");
		
		addRows = 0;
		deleteRows = 0;
		long begin = System.currentTimeMillis();
		int dateIndexTobeSort = 0;
		String outputPath = new String();
		if(testFlag) {
			outputPath = "/Users/rodrick.d/Desktop";
		}
		else {
			outputPath = updateEngine.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath();
			//decode to deal with Chinese characters and blanks
			outputPath = java.net.URLDecoder.decode(outputPath, "utf-8");
			outputPath = outputPath.replace("%20", " ");
			//acquire parent path
			outputPath = outputPath.substring(1, outputPath.lastIndexOf("/"));
		}

		File engineFile = new File(outputPath + "/Engine");
		if(!engineFile.exists()) {
			engineFile.mkdir();
		}
		//get date
		SimpleDateFormat sdf = new SimpleDateFormat("HHa_mm_ss MMM-dd-yyyy");	//set date format
		Calendar calendar = Calendar.getInstance();	
		Date date = calendar.getTime();	//get system date
		String dateStringPase = sdf.format(date);	//change date to string according to set format
	
		//output file name
		String outputFileName = "Powertrain Engine Tracking List_" + dateStringPase + ".xlsx";
		
		//open streams
		FileInputStream originFis = new FileInputStream(new File(originalPath));
		FileInputStream formatFis = new FileInputStream(new File(formatPath));
		FileOutputStream outputFos = new FileOutputStream(new File(outputPath + "/Engine/" + outputFileName));
		
		//open workbook
		XSSFWorkbook originWorkbook = new XSSFWorkbook(originFis);
		XSSFWorkbook formatWorkbook = new XSSFWorkbook(formatFis);
		XSSFWorkbook outputWorkbook = new XSSFWorkbook();
		
		//open sheet
		XSSFSheet originSheet = originWorkbook.getSheetAt(0);
		XSSFSheet formatSheet = formatWorkbook.getSheetAt(0);
		XSSFSheet outputSheet = outputWorkbook.createSheet("Updates");
		
		//set cell style
		CellStyle finishedCellStyle = outputWorkbook.createCellStyle();
		finishedCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		finishedCellStyle.setFillForegroundColor(IndexedColors.LIME.getIndex());
		
		CellStyle inProcessCellStyle = outputWorkbook.createCellStyle();
		inProcessCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		inProcessCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		
		CellStyle problemCellStyle = outputWorkbook.createCellStyle();
		problemCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		problemCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
		
		CellStyle titleCellStyle = outputWorkbook.createCellStyle();
		XSSFFont font = outputWorkbook.createFont();
		font.setColor((short)1);
		font.setBold(true);
		font.setFontHeight(12);
		titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
		titleCellStyle.setFont(font);
		titleCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		titleCellStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
		
		//copy title
		for(int i = 0; i < 6; i++) {
			XSSFRow outputRow = outputSheet.createRow(i);
			XSSFRow originRow = originSheet.getRow(i);
			int columnNum = 0;
			if(originRow != null) {
				columnNum = originRow.getLastCellNum();
			}
			for(int j = 0; j < columnNum; j++) {
				outputRow.createCell(j + 2).setCellValue(ExcelValue.getValue(originRow.getCell(j)));
			}
			switch (i) {
			case 1:
				outputRow.createCell(9).setCellStyle(finishedCellStyle);
				outputRow.createCell(10).setCellValue("Finished");
				break;
			case 2:
				outputRow.createCell(9).setCellStyle(inProcessCellStyle);
				outputRow.createCell(10).setCellValue("In Process");
				break;
			case 3:
				outputRow.createCell(9).setCellStyle(problemCellStyle);
				outputRow.createCell(10).setCellValue("Overdue/Problem");
				break;
			default:
				break;
			}
		}
		
		ArrayList<Integer> dateCellIndex = new ArrayList<Integer>();
		int planningObjIndex = 0;
		int purpofUseIndex = 0;
		//title row
		XSSFRow outputTitleRow = outputSheet.createRow(7);
		XSSFRow originTitleRow = originSheet.getRow(9);
		int titleColumnNum = originTitleRow.getLastCellNum();
		XSSFCell addnewCell_1 = outputTitleRow.createCell(0);
		XSSFCell addnewCell_2 = outputTitleRow.createCell(1);
		addnewCell_1.setCellValue("Updates");
		addnewCell_1.setCellStyle(titleCellStyle);
		addnewCell_2.setCellValue("Status");
		addnewCell_2.setCellStyle(titleCellStyle);
		for(int j = 0; j < titleColumnNum; j++) {
			XSSFCell outputTitleCell = outputTitleRow.createCell(j + 2);
			XSSFCell originTitleCell = originTitleRow.getCell(j);
			String titleName = ExcelValue.getValue(originTitleCell);
			outputTitleCell.setCellStyle(titleCellStyle);
			outputTitleCell.setCellValue(titleName);
			
			if(titleName.equals("Planning obj")) {
				planningObjIndex = (j + 2);
			}
			else if(titleName.equals("Purp. of use")) {
				purpofUseIndex = (j + 2);
			}
			else if(titleName.equals("Assembly Target - to")) {
				dateIndexTobeSort = j + 2;
				dateCellIndex.add(j + 2);
			}
			else if(titleName.equals("Assembly Current - from")) {
				dateCellIndex.add(j + 2);
			}
			else if(titleName.equals("Assembly Current - to")) {
				dateCellIndex.add(j + 2);
			}
			else if(titleName.equals("Requirement date")) {
				dateCellIndex.add(j + 2);
			}
		}
		
		//read and change
		ArrayList<Integer> problemProcess = new ArrayList<Integer>();
		ArrayList<Integer> InUseProcess = new ArrayList<Integer>();
		ArrayList<Integer> normalProcess = new ArrayList<Integer>();
		ArrayList<Integer> newNormalProcess = new ArrayList<Integer>();
		ArrayList<Integer> finishedProcess = new ArrayList<Integer>();
		
		Map<Integer, Integer> problemProcessMap = new HashMap<Integer, Integer>();
		Map<Integer, Integer> InUseProcessMap = new HashMap<Integer, Integer>();
		Map<Integer, Integer> normalProcessMap = new HashMap<Integer, Integer>();
		Map<Integer, Integer> finishedProcessMap = new HashMap<Integer, Integer>();
		
		int originRowNum = originSheet.getLastRowNum() + 1;
		int formatRowNum = formatSheet.getLastRowNum() + 1;
		
		flag: 
		for(int i = 10; i < originRowNum; i++) {
			XSSFRow readOriginRow = originSheet.getRow(i);
			if(readOriginRow == null) {
				continue;
			}
			XSSFCell readOriginCell = readOriginRow.getCell(0);
			String originProcessNumber = ExcelValue.getValue(readOriginCell);
			for(int j = 8; j < formatRowNum; j++) {
				XSSFRow readFormatRow = formatSheet.getRow(j);
				if(readFormatRow == null) {
					continue;
				}
				XSSFCell readFormatCell = readFormatRow.getCell(2);
				//if exists
				if(ExcelValue.getValue(readFormatCell).equals(originProcessNumber)) {
					String cellColor = cellOperation.getARGBHex(readFormatCell);
//					System.out.println(cellColor);
					if(cellColor.equals("FF0000")) {
						problemProcess.add(j);
						problemProcessMap.put(j, i);
//						problemProcess.add(i);
						continue flag;
					}
					else if(cellColor.equals("FFFF00")) {
						InUseProcess.add(j);
						InUseProcessMap.put(j, i);
//						InUseProcess.add(i);
						continue flag;
					}
					else if(cellColor.equals("92D050") || cellColor.equals("99CC00") || cellColor.equals("00B050") || cellColor.equals("70AD47")) {
						finishedProcess.add(j);
						finishedProcessMap.put(j, i);
//						finishedProcess.add(i);
						continue flag;
					}
					else {
						normalProcess.add(j);
						normalProcessMap.put(j, i);
//						normalProcess.add(i);
						continue flag;
					}
				}
			}
			newNormalProcess.add(i);
			addRows++;
		}
		
		//write in output sheet
		int problemProcessNum = problemProcess.size();
		int inUseprocessNum = InUseProcess.size();
		int finishedProcessNum = finishedProcess.size();
		int normalProcessNum = normalProcess.size() + newNormalProcess.size();
		int insertRows = problemProcessNum + inUseprocessNum + finishedProcessNum + normalProcessNum;
		
		//sort
//		ArrayList<Integer> realProblemProcess = new ArrayList<Integer>();
		
//		int[] sortedProblemProcess = Sort.insertSort(problemProcess, dateIndexTobeSort, originSheet, false);
//		int[] sortedInUseProcess = Sort.insertSort(InUseProcess, dateIndexTobeSort, originSheet, false);

		for(int i = 0; i < insertRows; i++) {
			XSSFRow newOutputRow = outputSheet.createRow(i + 8);
			if(i < problemProcessNum) {
				for(int j = 0; j < titleColumnNum + 2; j++) {
					XSSFCell newCell = newOutputRow.createCell(j);
					newCell.setCellStyle(problemCellStyle);
					if(j == 0 || j == 1) {
//						cellOperation.copyCell(formatSheet.getRow(sortedProblemProcess[i]).getCell(j), newCell);
						cellOperation.copyCell(formatSheet.getRow(problemProcess.get(i)).getCell(j), newCell);
					}
					else {
//						cellOperation.copyCell(originSheet.getRow(problemProcessMap.get(sortedProblemProcess[i])).getCell(j - 2), newCell);
						cellOperation.copyCell(originSheet.getRow(problemProcessMap.get(problemProcess.get(i))).getCell(j - 2), newCell);
					}
					
					if(dateCellIndex.contains(j)) {
						String newCellValue = ExcelValue.getValue(newCell);
						if(newCellValue != null && !newCellValue.equals("")) {
							newCell.setCellValue(newCellValue.substring(3, 6) + newCellValue.substring(0, 3) + newCellValue.substring(6));
						}
						cellOperation.setDateFormat(newCell, problemCellStyle);
					}
					
				}
			}
			else if(i < (problemProcessNum + inUseprocessNum)) {
				for(int j = 0; j < titleColumnNum + 2; j++) {
					XSSFCell newCell = newOutputRow.createCell(j);
					newCell.setCellStyle(inProcessCellStyle);
					if(j == 0 || j == 1) {
//						cellOperation.copyCell(formatSheet.getRow(sortedInUseProcess[(i - problemProcessNum)]).getCell(j), newCell);
						cellOperation.copyCell(formatSheet.getRow(InUseProcess.get((i - problemProcessNum))).getCell(j), newCell);
					}
					else {
//						cellOperation.copyCell(originSheet.getRow(InUseProcessMap.get(sortedInUseProcess[(i - problemProcessNum)])).getCell(j - 2), newCell);
						cellOperation.copyCell(originSheet.getRow(InUseProcessMap.get(InUseProcess.get((i - problemProcessNum)))).getCell(j - 2), newCell);
					}
					
					if(dateCellIndex.contains(j)) {
						String newCellValue = ExcelValue.getValue(newCell);
						if(newCellValue != null && !newCellValue.equals("")) {
							newCell.setCellValue(newCellValue.substring(3, 6) + newCellValue.substring(0, 3) + newCellValue.substring(6));
						}
						cellOperation.setDateFormat(newCell, inProcessCellStyle);
					}
					
				}
			}
			else if(i < (problemProcessNum + inUseprocessNum + normalProcessNum)) {
				if(i < (problemProcessNum + inUseprocessNum + normalProcess.size())) {
					for(int j = 0; j < titleColumnNum + 2; j++) {
						XSSFCell newCell = newOutputRow.createCell(j);
						if(j == 0 || j == 1) {
							cellOperation.copyCell(formatSheet.getRow(normalProcess.get(i - problemProcessNum - inUseprocessNum)).getCell(j), newCell);
						}
						else {
							cellOperation.copyCell(originSheet.getRow(normalProcessMap.get(normalProcess.get(i - problemProcessNum - inUseprocessNum))).getCell(j - 2), newCell);
						}
						
						if(dateCellIndex.contains(j)) {
							String newCellValue = ExcelValue.getValue(newCell);
							if(newCellValue != null && !newCellValue.equals("")) {
								newCell.setCellValue(newCellValue.substring(3, 6) + newCellValue.substring(0, 3) + newCellValue.substring(6));
							}
							cellOperation.setDateFormat(newCell, null);
						}
					}
				}
				else {
					newOutputRow.createCell(0);
					newOutputRow.createCell(1);
					for(int j = 0; j < titleColumnNum; j++) {
						XSSFCell newCell = newOutputRow.createCell(j + 2);
						cellOperation.copyCell(originSheet.getRow(newNormalProcess.get(i - problemProcessNum - inUseprocessNum - normalProcess.size())).getCell(j), newCell);
						if(dateCellIndex.contains(j + 2)) {
							String newCellValue = ExcelValue.getValue(newCell);
							if(newCellValue != null && !newCellValue.equals("")) {
								newCell.setCellValue(newCellValue.substring(3, 6) + newCellValue.substring(0, 3) + newCellValue.substring(6));
							}
							cellOperation.setDateFormat(newCell, null);
						}
					}
				}
			}
			else {
				for(int j = 0; j < titleColumnNum + 2; j++) {
					XSSFCell newCell = newOutputRow.createCell(j);
					newCell.setCellStyle(finishedCellStyle);
					if(j == 0 || j == 1) {
						cellOperation.copyCell(formatSheet.getRow(finishedProcess.get(i - problemProcessNum - inUseprocessNum - normalProcessNum)).getCell(j), newCell);
					}
					else {
						cellOperation.copyCell(originSheet.getRow(finishedProcessMap.get(finishedProcess.get(i - problemProcessNum - inUseprocessNum - normalProcessNum))).getCell(j - 2), newCell);
					}
					
					if(dateCellIndex.contains(j)) {
						String newCellValue = ExcelValue.getValue(newCell);
						if(newCellValue != null && !newCellValue.equals("")) {
							newCell.setCellValue(newCellValue.substring(3, 6) + newCellValue.substring(0, 3) + newCellValue.substring(6));
						}
						cellOperation.setDateFormat(newCell, finishedCellStyle);
					}
					
				}
			}
		}
		
		//sort the normal
		int normalStartIndex = 8 + problemProcessNum + inUseprocessNum;
		ArrayList<Integer> outputtedNormalProcess = new ArrayList<Integer>();
		for(int i = 0; i < normalProcessNum; i++) {
			outputtedNormalProcess.add(normalStartIndex + i);
		}
		int[] sortedNormalProcess = Sort.insertSort(outputtedNormalProcess, dateIndexTobeSort, outputSheet, true);
		String[][] normalContent = new String[normalProcessNum][titleColumnNum + 2];
		
		for(int i = 0; i < normalProcessNum; i++) {
			XSSFRow outputNormalRow = outputSheet.getRow(sortedNormalProcess[i]);
			for(int j = 0; j < titleColumnNum + 2; j++) {
				XSSFCell preCell = outputNormalRow.getCell(j);
				normalContent[i][j] = ExcelValue.getValue(preCell);
			}
		}
		
		for(int i = 0; i < normalProcessNum; i++) {
			XSSFRow preRow = outputSheet.getRow(i + normalStartIndex);
			for(int j = 0; j < titleColumnNum + 2; j++) {
				XSSFCell preCell = preRow.getCell(j);
				preCell.setCellValue(normalContent[i][j]);
				if(dateCellIndex.contains(j)) {
					cellOperation.setDateFormat(preCell, null);
				}
			}
		}
		
		//sort the problem
		int problemStartIndex = 8;
		ArrayList<Integer> outputtedProblemProcess = new ArrayList<Integer>();
		for(int i = 0; i < problemProcessNum; i++) {
			outputtedProblemProcess.add(problemStartIndex + i);
		}
		int[] sortedProblemProcess = Sort.insertSort(outputtedProblemProcess, dateIndexTobeSort, outputSheet, true);
		String[][] problemContent = new String[problemProcessNum][titleColumnNum + 2];
		
		for(int i = 0; i < problemProcessNum; i++) {
			XSSFRow outputProblemRow = outputSheet.getRow(sortedProblemProcess[i]);
			for(int j = 0; j < titleColumnNum + 2; j++) {
				XSSFCell preCell = outputProblemRow.getCell(j);
				problemContent[i][j] = ExcelValue.getValue(preCell);
			}
		}
		
		for(int i = 0; i < problemProcessNum; i++) {
			XSSFRow preRow = outputSheet.getRow(i + problemStartIndex);
			for(int j = 0; j < titleColumnNum + 2; j++) {
				XSSFCell preCell = preRow.getCell(j);
				preCell.setCellValue(problemContent[i][j]);
				if(dateCellIndex.contains(j)) {
					cellOperation.setDateFormat(preCell, problemCellStyle);
				}
			}
		}
		
		//sort the inuse
		int InUseStartIndex = 8 + problemProcessNum;
		ArrayList<Integer> outputtedInUseProcess = new ArrayList<Integer>();
		for(int i = 0; i < inUseprocessNum; i++) {
			outputtedInUseProcess.add(InUseStartIndex + i);
		}
		int[] sortedInUseProcess = Sort.insertSort(outputtedInUseProcess, dateIndexTobeSort, outputSheet, true);
		String[][] InUseContent = new String[inUseprocessNum][titleColumnNum + 2];
		
		for(int i = 0; i < inUseprocessNum; i++) {
			XSSFRow outputInUseRow = outputSheet.getRow(sortedInUseProcess[i]);
			for(int j = 0; j < titleColumnNum + 2; j++) {
				XSSFCell preCell = outputInUseRow.getCell(j);
				InUseContent[i][j] = ExcelValue.getValue(preCell);
			}
		}
		
		for(int i = 0; i < inUseprocessNum; i++) {
			XSSFRow preRow = outputSheet.getRow(i + InUseStartIndex);
			for(int j = 0; j < titleColumnNum + 2; j++) {
				XSSFCell preCell = preRow.getCell(j);
				preCell.setCellValue(InUseContent[i][j]);
				if(dateCellIndex.contains(j)) {
					cellOperation.setDateFormat(preCell, inProcessCellStyle);
				}
			}
		}

		//column width of date column
		for(int i = 0; i < dateCellIndex.size(); i++) {
			outputSheet.autoSizeColumn(dateCellIndex.get(i));
		}
		outputSheet.setColumnWidth(0, (int)(50 + 0.72) * 256);
		outputSheet.autoSizeColumn(1);
		outputSheet.autoSizeColumn(purpofUseIndex);
		outputSheet.autoSizeColumn(planningObjIndex);
		
		outputWorkbook.write(outputFos);
		outputFos.flush();
		if(outputWorkbook != null) {
			outputWorkbook.close();
		}
		if(originWorkbook != null) {
			originWorkbook.close();
		}
		if(formatWorkbook != null) {
			formatWorkbook.close();
		}
		if(originFis != null) {
			originFis.close();
		}
		if(formatFis != null) {
			formatFis.close();
		}
		if(outputFos != null) {
			outputFos.close();
		}
		deleteRows = (formatRowNum - 8) + addRows - (originRowNum - 10);
		long end = System.currentTimeMillis();
		
		System.out.println("Program finish! Time consuming: " + (end - begin) + " ms.");
		System.out.println(addRows + " rows have been added!");
		System.out.println(deleteRows + " rows have been deleteded!");
		
		return (outputPath + "/Engine/" + outputFileName);
	}
	
	public static void main(String[] args) {
		String originPath = "/Users/rodrick.d/Desktop/engine/Original.xlsx";
		String formatPath = "/Users/rodrick.d/Desktop/engine/Formatted.xlsx";
		try {
			updateEngine(originPath, formatPath);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
