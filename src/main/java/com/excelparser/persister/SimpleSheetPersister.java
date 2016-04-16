package com.excelparser.persister;

import com.excelparser.spreadsheet.SpreadSheet;

public class SimpleSheetPersister<K, V> implements SheetPersister<K, V> {

	private final Database<K, V> database;
	private int rowsPersisted;

	public SimpleSheetPersister(Database<K, V> database) {
		this.database = database;
	}

	@Override
	public void persist(final SpreadSheet<K, V> spreadSheet) {
		rowsPersisted += database.getSheetDao().create(spreadSheet);
	}

	@Override
	public int getRowsPersisted() {
		return rowsPersisted;
	}

}
