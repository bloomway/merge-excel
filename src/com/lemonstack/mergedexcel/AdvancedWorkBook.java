package com.lemonstack.mergedexcel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

public final class AdvancedWorkBook {

	private HSSFWorkbook workBook;
	private final SimpleDateFormat formatter = new SimpleDateFormat("dd-MM-yyyy");
	
	public AdvancedWorkBook() {
		this.workBook = new HSSFWorkbook();
	}
	
	/**
	 * Create a new sheet
	 * @param 
	 * 		sheetName name of the sheet
	 * @return 
	 * 		A new created sheet
	 */
	public final HSSFSheet createSheet(final String sheetName) {
		return this.workBook.createSheet(sheetName);
	}
	
	/**
	 * Get a sheet based on his name
	 * @param sheetName 
	 * 			name of the sheet
	 * @return 
	 * 		HSSFSheet with the name provided or null if it does not exist
	 */
	public final HSSFSheet getSheet(final String sheetName) {
		return this.workBook.getSheet(sheetName);
	}
	
	/**
	 * Add one sheet to another
	 * @param srcSheet 
	 * 		sheet from which data came from
	 * @param sheetName 
	 * 		name of the sheet
	 */
	public void addSheet(final HSSFSheet srcSheet, final String sheetName) {
		final HSSFSheet dstSheet = this.workBook.getSheet(sheetName);
		final int rowNum = dstSheet.getLastRowNum();
		
		for(int idx = srcSheet.getFirstRowNum(); idx<= srcSheet.getLastRowNum(); idx++) {
			final Row srcRow = srcSheet.getRow(idx);
			
			// if we copy another file other than the 1st file, we skip the header
			if (rowNum > 0) {
				// if the header (first line), skip it
				if (srcRow.getRowNum() == 0) {
					continue;
				}
			}
			
			final Row dstRow = dstSheet.createRow(rowNum + idx);
			if (null != srcRow) {
				copyRow(srcRow, dstRow);
			}
			
		}
		
		// apply the style from the old sheet to the new sheet
		final int maxColNum = srcSheet.getLastRowNum();
		for (int p = 0; p < maxColNum; p++) {
			dstSheet.setColumnWidth(p, srcSheet.getColumnWidth(p));
		}
	}
	
	/**
	 * Copy a row 
	 * @param srcRow source 
	 * @param dstRow destination
	 */
	private void copyRow(final Row srcRow, final Row dstRow) {
		for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {

			final Cell srcCell = srcRow.getCell(j);
			Cell dstCell = dstRow.getCell(j);

			if (null != srcCell) {
				if (dstCell == null) {
					dstCell = dstRow.createCell(j);
				}
				copyCell(srcCell, dstCell);
			}
		}
	}

	/**
	 * Copy a cell
	 * @param srcCell source
	 * @param dstCell destination
	 */
	private void copyCell(final Cell srcCell, final Cell dstCell) {
		
		switch (srcCell.getCellType()) {
		case HSSFCell.CELL_TYPE_FORMULA:
			dstCell.setCellFormula(srcCell.getCellFormula());
			break;
		case HSSFCell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(srcCell)) {
				final String dateFormatted = this.formatter.format(srcCell.getDateCellValue());
				dstCell.setCellValue(dateFormatted);
			} else {
				dstCell.setCellValue(srcCell.getNumericCellValue());
			}
			break;
		case HSSFCell.CELL_TYPE_BLANK:
			dstCell.setCellValue(srcCell.getStringCellValue());
			break;
		case HSSFCell.CELL_TYPE_BOOLEAN:
			dstCell.setCellValue(srcCell.getBooleanCellValue());
			break;
		case HSSFCell.CELL_TYPE_STRING:
			dstCell.setCellValue(srcCell.getStringCellValue());
			break;
		default:
			break;
		}	
	}

	public void writeTo(final File file) {
		OutputStream os = null;
		
		// create a destination file
		createFile(file);
		
		try {
			os = new FileOutputStream(file);
			this.workBook.write(os);
			this.workBook.close();
		} catch (FileNotFoundException e) {
			throw new RuntimeException(e.getMessage());
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage());
		} finally {
			if(null != os) {
				try {
					os.close();
				} catch (IOException e) {
					throw new RuntimeException(e.getMessage());
				}
			}
		}
	}

	/**
	 * Create a new file on the filesystem
	 * @param file
	 */
	private void createFile(final File file) {
		final File parent = file.getParentFile();
		
		try {
			if (!parent.exists()) {
				parent.mkdirs();
			}
			
			if (!file.exists()) {
				file.createNewFile();
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		
	}
}
