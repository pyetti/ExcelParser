package com.excelparser.persister;

import com.excelparser.spreadsheet.SpreadSheet;

public interface SheetDao<K, V> {

	int create(SpreadSheet<K, V> spreadSheet);

}
