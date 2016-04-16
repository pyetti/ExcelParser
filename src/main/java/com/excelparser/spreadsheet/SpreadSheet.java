package com.excelparser.spreadsheet;

import com.excelparser.matrixmap.Matrix;

public interface SpreadSheet<K, V> extends Matrix<K, V> {

	void setEndRowNum(int endRowNum);

	int getEndRowNum();

	void setStartRowNum(int startRowNum);

	int getStartRowNum();

	int getRowCount();

	void incrementRowCount();

}
