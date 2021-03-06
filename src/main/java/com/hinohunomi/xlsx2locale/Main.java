package com.hinohunomi.xlsx2locale;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Main {

	public static void main(String[] args) throws IOException, InvalidFormatException {
		File file = new File(args[0]);
		Workbook workbook = WorkbookFactory.create(file);
		Sheet sheet = workbook.getSheetAt(0);
		int colCount = getColCount(sheet);

		for (int i = 1; i < colCount; i++) {
			put(sheet, i);
		}
	}

	static int getColCount(Sheet sheet) {
		int result = 0;
		Row row = sheet.getRow(0);
		if (row != null) {
			result = row.getLastCellNum();
		}
		return result;
	}

	static void put(Sheet sheet, int col) throws IOException {
		Row row = sheet.getRow(0);
		if (row == null) return;
		Cell cell = row.getCell(col);
		if (cell == null) return;
		final String localeName = cell.getStringCellValue();

		int rowNum = sheet.getLastRowNum();
		List<String> list = new ArrayList<>(rowNum);
		
		for (int r = 0; r <= rowNum; r++) {
            row = sheet.getRow(r);
            if (row != null) {
    			String key = getCellString(row, 0);
    			if (StringUtils.isEmpty(key)) {
                	list.add("");
    			} else {
        			String v = getCellString(row, col);
                	list.add(key + "=" + v);
    			}
            } else {
            	list.add("");
            }
        }

		Path dirpath = Paths.get("locale", localeName);
		if (!Files.exists(dirpath)) {
			Files.createDirectory(dirpath);
		}
		Path path = dirpath.resolve("resources.properties");
		save(list, path);
	}

	static String getCellString(Row row, int cellNum) {
		Cell cell = row.getCell(cellNum);
		if (cell == null) return "";
//		switch (cell.getCellType()) {
//			case Cell.CELL_TYPE_NUMERIC:
//				cell.getNumericCellValue();
//				break;
//			case Cell.CELL_TYPE_STRING:
//				cell.getStringCellValue();
//				break;
//			default:
//				break;
//		}
		return cell.getStringCellValue();
	}

	static void save(List<String> list, Path path) throws IOException {
		Files.deleteIfExists(path);

		Files.write(path, list);
	}
}
