package com.lemonstack.mergedexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class MergeManager {

	public MergeManager() {
	}
	
	/**
	 * Merge a list of Excel files
	 * @param xlsFiles files to merge
	 * @param dstFile file merged
	 * @return file merged
	 */
	public File merge(final List<File> xlsFiles, final File dstFile) {
		final AdvancedWorkBook mergedBook = new AdvancedWorkBook();
		try {
			for (final File xlsFile : xlsFiles) {
				final HSSFWorkbook wbook = new HSSFWorkbook(new FileInputStream(xlsFile));
				final int numOfSheets = wbook.getNumberOfSheets();
				
				for (int i = 0; i < numOfSheets; i++) {
					
					// For example, if we're only interested in sheet 2 and sheet 3
					if (i == 2 || i == 3) {
						final HSSFSheet sheet = wbook.getSheetAt(i); // sheet i
						final String sheetName = sheet.getSheetName(); // name of sheet i

						// if it is 1st file, so mergedBook.getSheet(sheetName)
						// doesn't exist.
						if (mergedBook.getSheet(sheetName) == null) {
							mergedBook.createSheet(sheetName);
						}
						mergedBook.addSheet(sheet, sheetName);
					}
				}
				wbook.close();
			}
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage());
		}

		mergedBook.writeTo(dstFile);
		return dstFile;
	}
}
