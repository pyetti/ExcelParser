package com.excelparser.spreadsheet;

import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Predicate;

import com.excelparser.matrixmap.MatrixMap;

public class ExcelSpreadSheet implements MatrixMap<String, String> {

	private Map<String, String> excelMap;

	public ExcelSpreadSheet() {
		excelMap = new LinkedHashMap<String, String>();
	}

	@Override
	public List<String> getColumn(String column) {
		column.toUpperCase();
		List<String> columnList = new ArrayList<String>();
		ExcelConsumer excelConsumer = new ExcelConsumer(columnList, excelMap);
		ExcelPredicate predicate = new ExcelPredicate(column, (String s) -> s.substring(0, column.length()));
		excelMap.keySet().stream().filter(predicate)
				.forEachOrdered(excelConsumer);
		return columnList;
	}

	@Override
	public List<String> getColumn(String column, int start) {
		column.toUpperCase();
		List<String> columnList = new ArrayList<String>();
		ExcelConsumer excelConsumer = new ExcelConsumer(columnList, excelMap);
		ExcelPredicate predicate = new ExcelPredicate(column, (String s) -> s.substring(0, column.length()));
		excelMap.keySet().stream().skip(start).filter(predicate)
				.forEachOrdered(excelConsumer);
		return columnList;
	}

	@Override
	public List<String> getColumn(String column, int skip, int limit) {
		column.toUpperCase();
		List<String> columnList = new ArrayList<String>();
		ExcelConsumer excelConsumer = new ExcelConsumer(columnList, excelMap);
		ExcelPredicate predicate = new ExcelPredicate(column, (String s) -> s.substring(0, column.length()));
		excelMap.keySet().stream().filter(predicate).skip(skip).limit(limit)
				.forEachOrdered(excelConsumer);
		return columnList;
	}

	@Override
	public long getTotalColumnCells(String column) {
		column.toUpperCase();
		ExcelPredicate predicate = new ExcelPredicate(column, (String s) -> s.substring(0, column.length()));
		return excelMap.keySet().stream().filter(predicate).count();
	}

	@Override
	public List<String> getRow(int row) {
		List<String> rowList = new ArrayList<String>();
		ExcelPredicate predicate = new ExcelPredicate(String.valueOf(row), (String s) -> s.substring(1));
		ExcelConsumer excelConsumer = new ExcelConsumer(rowList, excelMap);
		excelMap.keySet().stream().filter(predicate)
				.forEachOrdered(excelConsumer);
		return rowList;
	}

	@Override
	public List<String> getRow(int row, int start) {
		List<String> rowList = new ArrayList<String>();
		ExcelPredicate predicate = new ExcelPredicate(String.valueOf(row), (String s) -> s.substring(1));
		ExcelConsumer excelConsumer = new ExcelConsumer(rowList, excelMap);
		excelMap.keySet().stream().skip(start).filter(predicate)
				.forEachOrdered(excelConsumer);
		return rowList;
	}

	@Override
	public List<String> getRow(int row, int skip, int limit) {
		List<String> rowList = new ArrayList<String>();
		ExcelPredicate predicate = new ExcelPredicate(String.valueOf(row), (String s) -> s.substring(1));
		ExcelConsumer excelConsumer = new ExcelConsumer(rowList, excelMap);
		excelMap.keySet().stream().filter(predicate).skip(skip).limit(limit)
				.forEachOrdered(excelConsumer);
		return rowList;
	}

	@Override
	public long getTotalRowCells(int row) {
		ExcelPredicate predicate = new ExcelPredicate(String.valueOf(row), (String s) -> s.substring(1));
		return excelMap.keySet().stream().filter(predicate).count();
	}

	@Override
	public String getEntry(String cell) {
		return excelMap.get(cell);
	}

	@Override
	public boolean clearEntry(String cell) {
		cell.toUpperCase();
		if (excelMap.containsKey(cell)) {
			excelMap.put(cell, "");
			return true;
		} else {
			return false;
		}
	}

	@Override
	public String put(String cell, String value) {
		cell.toUpperCase();
		return excelMap.put(cell, value);
	}

	@Override
	public int getTotalCells() {
		return excelMap.size();
	}

	@Override
	public boolean isEmpty() {
		return excelMap.isEmpty();
	}

	@Override
	public boolean containsCell(String cell) {
		cell.toUpperCase();
		return excelMap.containsKey(cell);
	}

	@Override
	public boolean containsEntry(String value) {
		return excelMap.containsValue(value);
	}

	@Override
	public Collection<String> getHeader() {
		return getRow(1);
	}

	@Override
	public String remove(String key) {
		return null;
	}

	@Override
	public void clear() {
		excelMap.clear();
	}

	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder();
		builder.append("Total Cells: ").append(getTotalCells());
		return builder.toString();
	}

	@FunctionalInterface
	public interface ExcelFunction {

		String subString(String s);

	}

	private class ExcelPredicate implements Predicate<String> {

		private final String column;
		private ExcelFunction excelFunction;

		public ExcelPredicate(String column, ExcelFunction f) {
			this.column = column;
			this.excelFunction = f;
		}

		@Override
		public boolean test(String s) {
			return column.equals(excelFunction.subString(s));
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
