package com.excelparser.persister;

import java.util.concurrent.Callable;

import com.excelparser.matrixmap.Matrix;

public class SheetPersistenceRunner<K, V> implements Callable<Integer> {

	private final Database<K, V> database;
	private final Matrix<K, V> spreadSheet;
	
	public SheetPersistenceRunner(final Database<K, V> database, final Matrix<K, V> spreadSheet) {
		this.database = database;
		this.spreadSheet = spreadSheet;
	}

	@Override
	public Integer call() throws Exception {
		return new SheetDaoImpl<K, V>(database.getDataSource(), database.getSql()).create(spreadSheet);
	}

}
