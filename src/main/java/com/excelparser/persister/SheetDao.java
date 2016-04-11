package com.excelparser.persister;

import com.excelparser.matrixmap.Matrix;

public interface SheetDao<K, V> {

	int create(Matrix<K, V> spreadSheet);

}
