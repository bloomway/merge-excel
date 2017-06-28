package com.lemonstack.mergedexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Merger {

	private static final String FILE_NAME = "N:\\workspace\\java-projects\\workspace\\merge-excel\\data\\merged.xls";

	public Merger() {
	}
	
	/**
	 * Merge files within a directory
	 * @param dir Directory containing the files to merge
	 */
	public void merge(final File dir) {
		final List<HSSFWorkbook> wbooks = getHSSFFiles(dir);
		final AdvancedWorkBook mergedBook = new AdvancedWorkBook(FILE_NAME);
		//HSSFWorkbook book = mergedBook.getWorkBook();
		
		//mergedBook.createSheet(AGENTS_NON_DECLARES);
		//mergedBook.createSheet(AGENTS_SUJETS_ALERTES);
		
		//final HSSFWorkbook wbook0 = wbooks.get(0);
		//mergedBook.addSheet(wbook0.getSheet(AGENTS_NON_DECLARES), AGENTS_NON_DECLARES);
		
		for(int k = 0, size = wbooks.size(); k < size; k++) {	
			final HSSFWorkbook wbook = wbooks.get(k); // file k (= 1, 2, 3, 4)	
			for(int i = 0, numOfSheets = wbook.getNumberOfSheets(); i < numOfSheets; i++) {
				
				// we're only interested in sheet 2 and sheet 3
				if (i == 2 || i == 3) {
					final HSSFSheet sheet = wbook.getSheetAt(i); // sheet i
					final String sheetName = sheet.getSheetName(); // name of sheet i

					// k == 0 => file 1
					// When it's the 1st file, we create a new sheet
					if (k == 0) {
						// i == 2 => sheet 1
						mergedBook.createSheet(sheetName);
					}
					mergedBook.addSheet(sheet, sheetName);
				}
				
			}
		}
		mergedBook.write();
	}
	
	/**
	 * 
	 * @param dir directory containing the files 
	 * 	we want to merge
	 * @return
	 */
	public List<HSSFWorkbook> getHSSFFiles(final File dir) {
		final List<HSSFWorkbook> wbookList = new ArrayList<>();
		
		final String[] files = dir.list();
		try {
			for (String fileName : files) {
				wbookList.add(new HSSFWorkbook(new FileInputStream(new File(dir, fileName))));
			}
		} catch (IOException e) {
			e.printStackTrace();
		}

		return wbookList;
	}

		   
}
