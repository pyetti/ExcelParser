package com.excelparser.spreadsheet;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Predicate;

import com.excelparser.matrixmap.Matrix;

public class ExcelSpreadSheet implements Matrix<String, String> {

	private Map<String, String> excelMap;

	public ExcelSpreadSheet() {
		excelMap = new LinkedHashMap<String, String>();
	}

	@Override
	public List<String> getCells(Predicate<String> predicate) {
		List<String> cells = new ArrayList<String>();
		ExcelConsumer<String, String> excelConsumer = new ExcelConsumer<String, String>(cells, excelMap);
		excelMap.values().stream().filter(predicate).forEach(excelConsumer);
		return cells;
	}

	@Override
	public List<String> getColumn(String column) {
		column.toUpperCase();
		List<String> columnList = new ArrayList<String>();
		ExcelConsumer<String, String> excelConsumer = new ExcelConsumer<String, String>(columnList, excelMap);
		excelMap.keySet().stream().filter((String  s) -> s.substring(0, column.length()).equals(column))
				.forEachOrdered(excelConsumer);
		return columnList;
	}

	@Override
	public List<String> getColumn(String column, int skip) {
		column.toUpperCase();
		List<String> columnList = new ArrayList<String>();
		ExcelConsumer<String, String> excelConsumer = new ExcelConsumer<String, String>(columnList, excelMap);
		excelMap.keySet().stream().skip(skip).filter((String  s) -> s.substring(0, column.length()).equals(column))
				.forEachOrdered(excelConsumer);
		return columnList;
	}

	@Override
	public List<String> getColumn(String column, int skip, int limit) {
		column.toUpperCase();
		List<String> columnList = new ArrayList<String>();
		ExcelConsumer<String, String> excelConsumer = new ExcelConsumer<String, String>(columnList, excelMap);
		excelMap.keySet().stream().filter((String  s) -> s.substring(0, column.length()).equals(column))
				.skip(skip).limit(limit)
				.forEachOrdered(excelConsumer);
		return columnList;
	}

	@Override
	public long getTotalColumnCells(String column) {
		column.toUpperCase();
		return excelMap.keySet().stream().filter((String  s) -> s.substring(0, column.length()).equals(column)).count();
	}

	@Override
	public List<String> getRow(int row) {
		List<String> rowList = new ArrayList<String>();
		ExcelConsumer<String, String> excelConsumer = new ExcelConsumer<String, String>(rowList, excelMap);
		excelMap.keySet().stream().filter((String s) -> s.substring(1).equals(String.valueOf(row)))
				.forEachOrdered(excelConsumer);
		return rowList;
	}

	@Override
	public List<String> getRow(int row, int skip) {
		List<String> rowList = new ArrayList<String>();
		ExcelConsumer<String, String> excelConsumer = new ExcelConsumer<String, String>(rowList, excelMap);
		excelMap.keySet().stream().skip(skip).filter((String s) -> s.substring(1).equals(String.valueOf(row)))
				.forEachOrdered(excelConsumer);
		return rowList;
	}

	@Override
	public List<String> getRow(int row, int skip, int limit) {
		List<String> rowList = new ArrayList<String>();
		ExcelConsumer<String, String> excelConsumer = new ExcelConsumer<String, String>(rowList, excelMap);
		excelMap.keySet().stream().filter((String s) -> s.substring(1).equals(String.valueOf(row)))
				.skip(skip).limit(limit)
				.forEachOrdered(excelConsumer);
		return rowList;
	}

	@Override
	public long getTotalRowCells(int row) {
		return excelMap.keySet().stream().filter((String s) -> s.substring(1).equals(String.valueOf(row))).count();
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
		return excelMap.put(cell.toUpperCase(), value);
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
	public boolean containsEntry(String value) {
		return excelMap.containsValue(value);
	}

	@Override
	public boolean containsCell(String cell) {
		cell.toUpperCase();
		return excelMap.containsKey(cell);
	}

	@Override
	public List<String> getHeader() {
		return getRow(1);
	}

	@Override
	public String removeColumn(String column, Move move) {
		if (Move.UP.equals(move)) {
			
		}
		return null;
	}

	@Override
	public String removeRow(String cell, Move move) {
		if (Move.UP.equals(move)) {
			
		}
		return null;
	}

	@Override
	public String removeCell(String cell, Move move) {
		if (Move.UP.equals(move)) {
			
		}
		return null;
	}

	@Override
	public void clear() {
		excelMap.clear();
	}

	@Override
	public Map<String, String> getSheet() {
		return excelMap;
	}

	@Override
	public Iterator<String> iterator() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder();
		builder.append("Total Cells: ").append(getTotalCells());
		return builder.toString();
	}

}
