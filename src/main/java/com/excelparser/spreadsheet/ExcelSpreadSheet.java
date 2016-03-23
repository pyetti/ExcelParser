package com.excelparser.spreadsheet;

import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import com.excelparser.matrixmap.MatrixMap;

public class ExcelSpreadSheet implements MatrixMap<String, String>  {

	private Map<String, String> excelMap;

	public ExcelSpreadSheet() {
		excelMap = new LinkedHashMap<String, String>();
	}

	@Override
	public List<String> getColumn(String column) {
		column.toUpperCase();
		List<String> columnList = new ArrayList<String>();
		ExcelConsumer<String, String>excelConsumer = new ExcelConsumer<String, String>(columnList, excelMap);
		ExcelPredicate<String, String, String> predicate = 
				new ExcelPredicate<String, String, String>(column, (String  s) -> s.substring(0, column.length()));
		excelMap.keySet().stream().filter(predicate)
				.forEachOrdered(excelConsumer);
		return columnList;
	}

	@Override
	public List<String> getColumn(String column, int skip) {
		column.toUpperCase();
		List<String> columnList = new ArrayList<String>();
		ExcelConsumer<String, String> excelConsumer = new ExcelConsumer<String, String>(columnList, excelMap);
		ExcelPredicate<String, String, String> predicate = 
				new ExcelPredicate<String, String, String>(column, (String s) -> s.substring(0, column.length()));
		excelMap.keySet().stream().skip(skip).filter(predicate)
				.forEachOrdered(excelConsumer);
		return columnList;
	}

	@Override
	public List<String> getColumn(String column, int skip, int limit) {
		column.toUpperCase();
		List<String> columnList = new ArrayList<String>();
		ExcelConsumer<String, String> excelConsumer = new ExcelConsumer<String, String>(columnList, excelMap);
		ExcelPredicate<String, String, String> predicate = 
				new ExcelPredicate<String, String, String>(column, (String s) -> s.substring(0, column.length()));
		excelMap.keySet().stream().filter(predicate).skip(skip).limit(limit)
				.forEachOrdered(excelConsumer);
		return columnList;
	}

	@Override
	public long getTotalColumnCells(String column) {
		column.toUpperCase();
		ExcelPredicate<String, String, String> predicate = 
				new ExcelPredicate<String, String, String>(column, (String s) -> s.substring(0, column.length()));
		return excelMap.keySet().stream().filter(predicate).count();
	}

	@Override
	public List<String> getRow(int row) {
		List<String> rowList = new ArrayList<String>();
		ExcelPredicate<String, String, String> predicate = 
				new ExcelPredicate<String, String, String>(String.valueOf(row), (String s) -> s.substring(1));
		ExcelConsumer<String, String> excelConsumer = new ExcelConsumer<String, String>(rowList, excelMap);
		excelMap.keySet().stream().filter(predicate)
				.forEachOrdered(excelConsumer);
		return rowList;
	}

	@Override
	public List<String> getRow(int row, int skip) {
		List<String> rowList = new ArrayList<String>();
		ExcelPredicate<String, String, String> predicate = 
				new ExcelPredicate<String, String, String>(String.valueOf(row), (String s) -> s.substring(1));
		ExcelConsumer<String, String> excelConsumer = new ExcelConsumer<String, String>(rowList, excelMap);
		excelMap.keySet().stream().skip(skip).filter(predicate)
				.forEachOrdered(excelConsumer);
		return rowList;
	}

	@Override
	public List<String> getRow(int row, int skip, int limit) {
		List<String> rowList = new ArrayList<String>();
		ExcelPredicate<String, String, String> predicate = 
				new ExcelPredicate<String, String, String>(String.valueOf(row), (String s) -> s.substring(1));
		ExcelConsumer<String, String> excelConsumer = new ExcelConsumer<String, String>(rowList, excelMap);
		excelMap.keySet().stream().filter(predicate).skip(skip).limit(limit)
				.forEachOrdered(excelConsumer);
		return rowList;
	}

	@Override
	public long getTotalRowCells(int row) {
		ExcelPredicate<String, String, String> predicate = 
				new ExcelPredicate<String, String, String>(String.valueOf(row), (String s) -> s.substring(1));
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
	public String remove(String key, Move move) {
		if (Move.UP.equals(move)) {
			
		}
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

}
