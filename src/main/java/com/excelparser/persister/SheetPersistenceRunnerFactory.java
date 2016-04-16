package com.excelparser.persister;

import com.excelparser.spreadsheet.SpreadSheet;

public class SheetPersistenceRunnerFactory<K, V> {

	public SheetPersistenceRunner<K, V> getSheetPersistenceRunner(final Database<K, V> database, 
			final SpreadSheet<K, V> spreadSheet) {
		return new SheetPersistenceRunner<K, V>(database, spreadSheet);
	}

}
