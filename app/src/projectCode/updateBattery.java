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

public class updateBattery {
	public static int addRows = 0;
	public static int deleteRows = 0;
	private static boolean testFlag = false;
	public static String updateBattery(String originalPath, String formatPath) throws Exception{
		//change the file path format
		originalPath = originalPath.replace("\\", "/");
		formatPath = formatPath.replace("\\", "/");
		
		addRows = 0;
		deleteRows = 0;
		long begin = System.currentTimeMillis();
		int dateIndexTobeSort = 0;
		int codeIndex = 0;
		int planningObjIndex = 0;
		int purpofUseIndex = 0;
		int placeofDelIndex = 0;
		//output path
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
		
		File batteryFile = new File(outputPath + "/Battery");
		if(!batteryFile.exists()) {
			batteryFile.mkdir();
		}
		//get date
		SimpleDateFormat sdf = new SimpleDateFormat("HHa_mm_ss MMM dd yyyy");	//set date format
		Calendar calendar = Calendar.getInstance();	
		Date date = calendar.getTime();	//get system date
		String dateStringPase = sdf.format(date);	//change date to string according to set format
	
		//output file name
		String outputFileName = "Powertrain Battery Tracking List_" + dateStringPase + ".xlsx";
		
		//code file
		String codeFilePath = new String();
		if(testFlag) {
			codeFilePath = "C:\\Users\\MINGDON\\Desktop\\Fangtong's task\\Battery code.xlsx";
		}
		else {
			codeFilePath = outputPath + "/Document/Battery code.xlsx";
		}	
		
		//open streams
		FileInputStream originFis = new FileInputStream(new File(originalPath));
		FileInputStream formatFis = new FileInputStream(new File(formatPath));
		FileInputStream codeFis = new FileInputStream(new File(codeFilePath));
		FileOutputStream outputFos = new FileOutputStream(new File(outputPath + "/Battery/" + outputFileName));
		
		//open workbook
		XSSFWorkbook originWorkbook = new XSSFWorkbook(originFis);
		XSSFWorkbook formatWorkbook = new XSSFWorkbook(formatFis);
		XSSFWorkbook codeWorkbook = new XSSFWorkbook(codeFis);
		XSSFWorkbook outputWorkbook = new XSSFWorkbook();
		
		//open sheet
		XSSFSheet originSheet = originWorkbook.getSheetAt(0);
		XSSFSheet formatSheet = formatWorkbook.getSheetAt(0);
		XSSFSheet codeSheet = codeWorkbook.getSheetAt(0);
		XSSFSheet outputSheet = outputWorkbook.createSheet("Updates");
		
		//code list
		int codeFileRowNum = codeSheet.getLastRowNum() + 1;
		String[][] codeList = new String[codeFileRowNum][2];
		for(int i = 0; i < codeFileRowNum; i++) {
			XSSFRow codeRow = codeSheet.getRow(i);
			codeList[i][0] = ExcelValue.getValue(codeRow.getCell(0));
			codeList[i][1] = ExcelValue.getValue(codeRow.getCell(1));
		}
			
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
		
		//title row
		XSSFRow outputTitleRow = outputSheet.createRow(7);
		XSSFRow originTitleRow = originSheet.getRow(9);
		int titleColumnNum = originTitleRow.getLastCellNum();
		XSSFCell addnewCell_1 = outputTitleRow.createCell(0);
		XSSFCell addnewCell_2 = outputTitleRow.createCell(1);
		addnewCell_1.setCellValue("Status");
		addnewCell_1.setCellStyle(titleCellStyle);
		addnewCell_2.setCellValue("Dummy");
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
			else if(titleName.equals("Place of del.")) {
				placeofDelIndex = (j + 2);
			}
			else if(titleName.equals("Requirement date")) {
				dateIndexTobeSort = j + 2;
			}
			else if(titleName.equals("HV-Battery development code")) {
				codeIndex = j + 2;
			}
			
		}
		
		//read and change
		ArrayList<Integer> problemProcess = new ArrayList<Integer>();
		ArrayList<Integer> InUseProcess = new ArrayList<Integer>();
		ArrayList<Integer> normalProcess = new ArrayList<Integer>();
		ArrayList<Integer> newNormalProcess = new ArrayList<Integer>();
		ArrayList<Integer> finishedProcess = new ArrayList<Integer>();
		int originRowNum = originSheet.getLastRowNum() + 1;
		int formatRowNum = formatSheet.getLastRowNum() + 1;
		
		Map<Integer, Integer> problemProcessMap = new HashMap<Integer, Integer>();
		Map<Integer, Integer> InUseProcessMap = new HashMap<Integer, Integer>();
		Map<Integer, Integer> normalProcessMap = new HashMap<Integer, Integer>();
		Map<Integer, Integer> finishedProcessMap = new HashMap<Integer, Integer>();
		
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
						continue flag;
					}
					else if(cellColor.equals("FFFF00")) {
						InUseProcess.add(j);
						InUseProcessMap.put(j, i);
						continue flag;
					}
					else if(cellColor.equals("92D050")|| cellColor.equals("99CC00")) {
						finishedProcess.add(j);
						finishedProcessMap.put(j, i);
						continue flag;
					}
					else {
						normalProcess.add(j);
						normalProcessMap.put(j, i);
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
//		int[] sortedProblemProcess = Sort.insertSort(problemProcess, dateIndexTobeSort, originSheet);
//		int[] sortedInUseProcess = Sort.insertSort(InUseProcess, dateIndexTobeSort, originSheet);

		for(int i = 0; i < insertRows; i++) {
			XSSFRow newOutputRow = outputSheet.createRow(i + 8);
			if(i < problemProcessNum) {
				for(int j = 0; j < titleColumnNum + 2; j++) {
					XSSFCell newCell = newOutputRow.createCell(j);
					newCell.setCellStyle(problemCellStyle);
					if(j == 0 || j == 1) {
						cellOperation.copyCell(formatSheet.getRow(problemProcess.get(i)).getCell(j), newCell);
					}
					else {
						cellOperation.copyCell(originSheet.getRow(problemProcessMap.get(problemProcess.get(i))).getCell(j - 2), newCell);
					}
					
					if(j == dateIndexTobeSort) {
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
						cellOperation.copyCell(formatSheet.getRow(InUseProcess.get((i - problemProcessNum))).getCell(j), newCell);
					}
					else {
						cellOperation.copyCell(originSheet.getRow(InUseProcessMap.get(InUseProcess.get((i - problemProcessNum)))).getCell(j - 2), newCell);
					}
					
					if(j == dateIndexTobeSort) {
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
						
						if(j == dateIndexTobeSort) {
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
						if((j + 2) == dateIndexTobeSort) {
							String newCellValue = ExcelValue.getValue(newCell);
							if(newCellValue != null && !newCellValue.equals("")) {
								newCell.setCellValue(newCellValue.substring(3, 6) + newCellValue.substring(0, 3) + newCellValue.substring(6));
							}
							cellOperation.setDateFormat(newCell, null);
						}
						else if((j + 2) == codeIndex) {
							String codeString = ExcelValue.getValue(newCell);
							String codes[] = codeString.split(";");
							OUT:
							for(String code : codes) {
								for(int k = 0; k < codeFileRowNum; k++) {
									if(codeList[k][0].equals(code)) {
										XSSFCell dummyCell = newCell.getRow().getCell(1);
										dummyCell.setCellValue(codeList[k][1]);
										break OUT;
									}
								}
							}
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
					
					if(j == dateIndexTobeSort) {
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
				if(j == dateIndexTobeSort) {
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
				if(j == dateIndexTobeSort) {
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
				if(j == dateIndexTobeSort) {
					cellOperation.setDateFormat(preCell, inProcessCellStyle);
				}
			}
		}
		
		//column width of date column
		outputSheet.autoSizeColumn(dateIndexTobeSort);
		outputSheet.setColumnWidth(0, (int)(50 + 0.72) * 256);
		outputSheet.autoSizeColumn(1);
		outputSheet.autoSizeColumn(placeofDelIndex);
		outputSheet.autoSizeColumn(planningObjIndex);
		outputSheet.autoSizeColumn(purpofUseIndex);
		
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
		if(codeWorkbook != null) {
			codeWorkbook.close();
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
	
		return (outputPath + "/Battery/" + outputFileName);
	}
	
	public static void main(String[] args) {
		String originPath = "C:\\Users\\MINGDON\\Desktop\\Fangtong's task\\Document\\Original_Battery.xlsx";
		String formatPath = "C:\\Users\\MINGDON\\Desktop\\Fangtong's task\\Document\\Format_Battery.xlsx";
		try {
			updateBattery(originPath, formatPath);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
