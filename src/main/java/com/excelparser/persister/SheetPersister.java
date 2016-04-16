package com.excelparser.persister;

import com.excelparser.spreadsheet.SpreadSheet;

public interface SheetPersister<K, V> {

	void persist(SpreadSheet<K, V> spreadSheet);
	int getRowsPersisted();

}
