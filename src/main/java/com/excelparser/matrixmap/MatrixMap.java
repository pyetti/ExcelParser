package com.excelparser.matrixmap;

import java.util.List;
import java.util.Map;

public interface MatrixMap<K, V> extends Map<K, V> {

	List<String> getColumn(K column);

	List<String> getColumn(K column, int start, int end);

	List<String> getColumn(K column, int start);

	long getColumnLength(K column);

	List<String> getRow(int row);

	List<String> getRow(int row, int start, int end);

	List<String> getRow(int row, int start);

	long getRowLength(int row);

	String getEntry(K cell);

}
