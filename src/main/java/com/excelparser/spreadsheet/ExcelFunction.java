package com.excelparser.spreadsheet;

@FunctionalInterface
public interface ExcelFunction<T, R> {
	R function(T t);
}
