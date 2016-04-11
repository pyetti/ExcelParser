package com.excelparser.persister;

import com.excelparser.matrixmap.Matrix;

public class SheetPersistenceRunnerFactory<K, V> {

	public SheetPersistenceRunner<K, V> getSheetPersistenceRunner(final Database<K, V> database, 
			final Matrix<K, V> spreadSheet) {
		return new SheetPersistenceRunner<K, V>(database, spreadSheet);
	}

}
