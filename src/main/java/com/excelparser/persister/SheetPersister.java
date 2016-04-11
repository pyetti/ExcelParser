package com.excelparser.persister;

import com.excelparser.matrixmap.Matrix;

public interface SheetPersister<K, V> {

	void persist(Matrix<K, V> spreadSheet);
	int getRowsPersisted();

}
