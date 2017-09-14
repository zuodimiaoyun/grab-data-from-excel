package com.pollos.tools.gdfe.util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	private static String exportName = "a";
	public static void grab(String path, String cellPositionStr) throws Exception {
		File dirctory = new File(path);
		if(!dirctory.isDirectory()){
			throw new Exception("input path not a directory!");
		}
		List<File> fileList = getListAndCanHandleFiles(dirctory, exportName);
		boolean isXlsx = isXlsx(fileList);
		String ext = isXlsx ? ".xlsx" : ".xls";
		if(fileList == null || fileList.isEmpty()){
			throw new Exception("directory is empty!");
		}
		String[] cellPositions = cellPositionStr.split(",");
		String exportFileName = exportName + ext;
		String pathDelimiter = OSInfo.isWindows() ? "\\" : "/";
		File exportFile = new File(path + pathDelimiter + exportFileName);
		Workbook exportBook = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook();
		createSheetsFromFile(exportFileName, exportBook, fileList, cellPositions);
		for (File file : fileList) {
			getFilterRowsFromFile(exportBook, file, cellPositions);
		}
		FileOutputStream out = new FileOutputStream(exportFile);
		exportBook.write(out);
		out.close();
		exportBook.close();
	}

	private static List<File> getListAndCanHandleFiles(File dirctory, String exportName) {
		List<File> fileList = new ArrayList<File>();
		File[] files = dirctory.listFiles(new FilenameFilter() {
			
			@Override
			public boolean accept(File dir, String name) {
				if((name.endsWith(".xls") || name.endsWith(".xlsx")) && !name.split("\\.")[0].equals(exportName)){
					return true;
				}
				return false;
			}
		});
		for (File file : files) {
			fileList.add(file);
		}
		Collections.sort(fileList, new Comparator<File>(){
			@Override
			public int compare(File o1, File o2) {
				String name1 = o1.getName();
				String name2 = o2.getName();
				if(name1.length() != name2.length()){
					return name1.length() - name2.length();
				}else{
					return name1.compareTo(name2);
				}
			}
			
		});
		return fileList;
	}

	private static boolean isXlsx(List<File> files) {
		for (File file : files) {
			if(file.getName().endsWith(".xls")){
				return false;
			}else if(file.getName().endsWith(".xlsx")){
				return true;
			}
			continue;
		}
		return false;
	}
	private static boolean isExcel(File file) {
		if(file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx")){
			return true;
		}
		return false;
	}

	private static void createSheetsFromFile(String exportFileName, Workbook exportFile, List<File> files, String[] cellPositions) throws FileNotFoundException, IOException, EncryptedDocumentException, InvalidFormatException {
		for (File file : files) {
			if(isExcel(file)){
				int sheetNum = getSheetNumFromExcel(file);
				for(int sheetIndex = 0; sheetIndex < sheetNum; sheetIndex ++){
					Sheet sheet = exportFile.createSheet();
					Row firstRow = sheet.createRow(0);
					firstRow.createCell(0).setCellValue("±í-Ò³Ç©");
					int index = 1;
					for (String cellPos : cellPositions) {
						firstRow.createCell(index).setCellValue(cellPos);
						index ++;
					}
				}
				break;
			}
		}
	}

	private static int getSheetNumFromExcel(File file) throws FileNotFoundException, IOException, EncryptedDocumentException, InvalidFormatException {
		Workbook wb = WorkbookFactory.create(file);
		return wb.getNumberOfSheets();
	}

	private static void getFilterRowsFromFile(Workbook exportBook, File file, String[] cellPositions) throws FileNotFoundException, IOException, EncryptedDocumentException, InvalidFormatException {
		
		Workbook wb = WorkbookFactory.create(file);
		for(int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex ++){
			Sheet sheet = wb.getSheetAt(sheetIndex);
			Sheet exoprtSheet = exportBook.getSheetAt(sheetIndex);
			Row exportRow = exoprtSheet.createRow(exoprtSheet.getLastRowNum() + 1);
			exportRow.createCell(0).setCellValue(file.getName() + "-Sheet" + sheetIndex);
			int cellIndex = 1;
			for(String cellPos: cellPositions){
				CellAddress cellAddr = new CellAddress(cellPos);
				Row row = sheet.getRow(cellAddr.getRow());
				Cell cell = row.getCell(cellAddr.getColumn());
				Cell exportCell = exportRow.createCell(cellIndex);
				copyCell(exportBook, exportCell, cell);
				cellIndex ++;
			}
		}
	}
	
	private static void copyCell(Workbook workbook, Cell newCell, Cell oldCell){
	    // Copy style from old cell and apply to new cell
//		CellStyle newCellStyle = workbook.createCellStyle();
//        newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
//        newCell.setCellStyle(newCellStyle);

        // If there is a cell comment, copy
        if (oldCell.getCellComment() != null) {
            newCell.setCellComment(oldCell.getCellComment());
        }

        // If there is a cell hyperlink, copy
        if (oldCell.getHyperlink() != null) {
            newCell.setHyperlink(oldCell.getHyperlink());
        }

        // Set the cell data type
        newCell.setCellType(oldCell.getCellTypeEnum());
        
        // Set the cell data value
        switch (oldCell.getCellTypeEnum()) {
            case BLANK:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case STRING:
                newCell.setCellValue(oldCell.getRichStringCellValue());
                break;
            case _NONE:
            	break;
            default:
            	break;
        }
	}
}
