package com.excelparser.matrixmap;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Predicate;

public class ExcelMatrixMap implements MatrixMap {

	private Map<String, String> excelMap;

	public ExcelMatrixMap() {
		excelMap = new LinkedHashMap<String, String>();
	}

	@Override
	public List<String> getColumn(String column) {
		List<String> columnList = new ArrayList<String>();
		ExcelConsumer excelConsumer = new ExcelConsumer(columnList, excelMap);
		ColumnPredicate predicate = new ColumnPredicate(column);
		excelMap.keySet().stream()
			.filter(predicate)
			.forEachOrdered(excelConsumer);
		return columnList;
	}

	@Override
	public List<String> getColumn(String column, int start) {
		List<String> columnList = new ArrayList<String>();
		ExcelConsumer excelConsumer = new ExcelConsumer(columnList, excelMap);
		ColumnPredicate predicate = new ColumnPredicate(column);
		excelMap.keySet().stream().skip(start)
			.filter(predicate)
			.forEachOrdered(excelConsumer);
		return columnList;
	}

	@Override
	public List<String> getColumn(String column, int skip, int limit) {
		List<String> columnList = new ArrayList<String>();
		ExcelConsumer excelConsumer = new ExcelConsumer(columnList, excelMap);
		ColumnPredicate predicate = new ColumnPredicate(column);
		excelMap.keySet().stream()
			.filter(predicate).skip(skip).limit(limit)
			.forEachOrdered(excelConsumer);
		return columnList;
	}

	@Override
	public long getColumnLength(String column) {
		ColumnPredicate predicate = new ColumnPredicate(column);
		return excelMap.keySet().stream()
			.filter(predicate).count();
	}

	@Override
	public List<String> getRow(int row) {
		List<String> rowList = new ArrayList<String>();
		RowPredicate predicate = new RowPredicate(String.valueOf(row));
		ExcelConsumer excelConsumer = new ExcelConsumer(rowList, excelMap);
		excelMap.keySet().stream()
			.filter(predicate)
			.forEachOrdered(excelConsumer);
		return rowList;
	}

	@Override
	public List<String> getRow(int row, int start) {
		List<String> rowList = new ArrayList<String>();
		RowPredicate predicate = new RowPredicate(String.valueOf(row));
		ExcelConsumer excelConsumer = new ExcelConsumer(rowList, excelMap);
		excelMap.keySet().stream().skip(start)
			.filter(predicate)
			.forEachOrdered(excelConsumer);
		return rowList;
	}

	@Override
	public List<String> getRow(int row, int skip, int limit) {
		List<String> rowList = new ArrayList<String>();
		RowPredicate predicate = new RowPredicate(String.valueOf(row));
		ExcelConsumer excelConsumer = new ExcelConsumer(rowList, excelMap);
		excelMap.keySet().stream()
			.filter(predicate).skip(skip).limit(limit)
			.forEachOrdered(excelConsumer);
		return rowList;
	}

	@Override
	public long getRowLength(int row) {
		RowPredicate predicate = new RowPredicate(String.valueOf(row));
		return excelMap.keySet().stream()
			.filter(predicate).count();
	}

	@Override
	public String getCellValue(String cell) {
		return excelMap.get(cell);
	}

	@Override
	public void put(String cell, String value) {
		this.excelMap.put(cell, value);
	}
	
	private class ColumnPredicate implements Predicate<String> {
		
		private final String column;
		
		public ColumnPredicate(String column) {
			this.column = column;
		}

		@Override
		public boolean test(String s) {
			return column.equals(s.substring(0, column.length()));
		}
		
	}
	
	private class RowPredicate implements Predicate<String> {
		
		private final String row;

		public RowPredicate(String row) {
			this.row = row;
		}

		@Override
		public boolean test(String s) {
			return row.equals(s.substring(1));
		}
		
	}
	
	private class ExcelConsumer implements Consumer<String> {
		
		private final List<String> list;
		private final Map<String, String> excelMap;
		
		public ExcelConsumer(List<String> list, Map<String, String> excelMap) {
			this.list = list;
			this.excelMap = excelMap;
		}
		
		@Override
		public void accept(String s) {
			list.add(excelMap.get(s));
		}
		
	}

}
