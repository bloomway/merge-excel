package com.lemonstack.mergedexcel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class AdvancedWorkBook {

	private HSSFWorkbook workBook;
	private String fileName;
	private int rowNum;
	
	public AdvancedWorkBook(String fileName) {
		this.workBook = new HSSFWorkbook();
		this.fileName = fileName;
	}
	
	public HSSFSheet createSheet(final String sheetName) {
		return this.workBook.createSheet(sheetName);
	}
	
	public void addSheet(final HSSFSheet srcSheet, final String sheetName) {
		
		final HSSFSheet dstSheet = this.workBook.getSheet(sheetName);
		
		for(int i = srcSheet.getFirstRowNum(); i<= srcSheet.getLastRowNum(); i++) {
			
			final Row srcRow = srcSheet.getRow(i);
			
			// if we copy another file other than the 1st file, we skip the header
			if (this.rowNum > 0) {
				if (srcRow.getRowNum() == 0) {
					continue;
				}
			}
			
			Row dstRow = dstSheet.createRow(this.rowNum++);
			if (null != srcRow) {
				copyRow(srcRow, dstRow);
				this.rowNum++;
			}
			
		}
		
		// apply the style from the old sheet to the new sheet
		final int maxColNum = srcSheet.getLastRowNum();
		for (int p = 0; p < maxColNum; p++) {
			dstSheet.setColumnWidth(p, srcSheet.getColumnWidth(p));
		}
	}
	
	
	private void copyRow(Row srcRow, Row dstRow) {
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

	private void copyCell(Cell srcCell, Cell dstCell) {
		
		switch (srcCell.getCellType()) {
		case HSSFCell.CELL_TYPE_FORMULA:
			dstCell.setCellFormula(srcCell.getCellFormula());
			break;
		case HSSFCell.CELL_TYPE_NUMERIC:
			dstCell.setCellValue(srcCell.getNumericCellValue());
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
			dstCell.setCellValue(srcCell.getDateCellValue());
			break;
		}	
	}

	public void write() {
		OutputStream os = null;
		try {
			os = new FileOutputStream(new File(fileName));
			this.workBook.write(os);
			this.workBook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if(null != os) {
				try {
					os.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
}
