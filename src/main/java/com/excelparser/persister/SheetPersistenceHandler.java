package com.excelparser.persister;

import org.apache.poi.xssf.model.SharedStringsTable;

import com.excelparser.parser.SheetHandler;
import com.excelparser.spreadsheet.ExcelSpreadSheet;
import com.excelparser.spreadsheet.SpreadSheet;

public class SheetPersistenceHandler extends SheetHandler {

	private SheetPersister<String, String> sheetPersister;
	private int numColumnsExpected;
	private int numColumnsRead;
	private int row;
	private boolean newSheet;
	public static final int MAX_ROWS_PER_READ = 1000;
	private SpreadSheet<String, String> spreadSheet;

	public SheetPersistenceHandler(SharedStringsTable sst,
			SheetPersister<String, String> sheetPersister, 
			SpreadSheet<String, String> spreadSheet, 
			int numColumnsExpected) {
		super(sst);
		this.sheetPersister = sheetPersister;
		this.numColumnsExpected = numColumnsExpected;
		this.spreadSheet = spreadSheet;
	}
	
	public SheetPersistenceHandler(SharedStringsTable sst,
			SheetPersister<String, String> sheetPersister, 
			int numColumnsExpected) {
		super(sst);
		this.sheetPersister = sheetPersister;
		this.numColumnsExpected = numColumnsExpected;
		this.spreadSheet = new ExcelSpreadSheet();
	}

	@Override
	protected void handleNewRow() {
		row++;
		if (row > 1) {
			numColumnsRead = 0;
			if (newSheet) {
				spreadSheet.setStartRowNum(row);
				spreadSheet.setEndRowNum(row);
				spreadSheet.incrementRowCount();
				newSheet = false;
			} else {
				spreadSheet.setEndRowNum(row);
				spreadSheet.incrementRowCount();
			}
		} 
	}

	@Override
	protected void handleCellData()  {
		if (row > 1) {
			spreadSheet.put(getCell(), getCellValue());
			numColumnsRead++;
			if (numColumnsRead % numColumnsExpected == 0 && row % MAX_ROWS_PER_READ == 0) {
				sheetPersister.persist(spreadSheet);
				spreadSheet = new ExcelSpreadSheet();
				newSheet = true;
				numColumnsRead = 0;
			}
		}
	}

	@Override
	public SpreadSheet<String, String> getSpreadSheet() {
		return spreadSheet;
	}

}
