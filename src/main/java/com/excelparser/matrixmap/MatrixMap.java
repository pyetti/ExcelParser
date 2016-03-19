package com.excelparser.matrixmap;

import java.util.List;

public interface MatrixMap {

	List<String> getColumn(String column);

	List<String> getRow(int row);

	String getCellValue(String cell);

	void put(String cell, String value);

	List<String> getRow(int row, int start, int end);

	List<String> getRow(int row, int start);

	List<String> getColumn(String column, int start, int end);

	List<String> getColumn(String column, int start);

	long getRowLength(int row);

	long getColumnLength(String column);

}
