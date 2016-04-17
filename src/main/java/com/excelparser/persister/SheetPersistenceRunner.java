package com.excelparser.persister;

import java.util.concurrent.Callable;

import com.excelparser.spreadsheet.SpreadSheet;

public class SheetPersistenceRunner<K, V> implements Callable<Integer> {

	private final Database<K, V> database;
	private final SpreadSheet<K, V> spreadSheet;
	
	public SheetPersistenceRunner(final Database<K, V> database, final SpreadSheet<K, V> spreadSheet) {
		this.database = database;
		this.spreadSheet = spreadSheet;
	}

	@Override
	public Integer call() throws Exception {
		return new SheetDaoImpl<K, V>(database).create(spreadSheet);
	}

}
