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
		excelMap.values().stream().filter(predicate).forEach(s -> cells.add(excelMap.get(s)));
		return cells;
	}

	@Override
	public List<String> getColumn(String column) {
		column.toUpperCase();
		List<String> columnList = new ArrayList<String>();
		excelMap.keySet().stream().filter(s -> s.substring(0, column.length()).equals(column))
				.forEachOrdered(s -> columnList.add(excelMap.get(s)));
		return columnList;
	}

	@Override
	public List<String> getColumn(String column, int skip) {
		column.toUpperCase();
		List<String> columnList = new ArrayList<String>();
		excelMap.keySet().stream().filter(s -> s.substring(0, column.length()).equals(column))
				.skip(skip)
				.forEachOrdered(s -> columnList.add(excelMap.get(s)));
		return columnList;
	}

	@Override
	public List<String> getColumn(String column, int skip, int limit) {
		column.toUpperCase();
		List<String> columnList = new ArrayList<String>();
		excelMap.keySet().stream().filter(s -> s.substring(0, column.length()).equals(column))
				.skip(skip).limit(limit)
				.forEachOrdered(s -> columnList.add(excelMap.get(s)));
		return columnList;
	}

	@Override
	public long getTotalColumnCells(String column) {
		column.toUpperCase();
		return excelMap.keySet().stream().filter(s -> s.substring(0, column.length()).equals(column)).count();
	}

	@Override
	public List<String> getRow(int row) {
		List<String> rowList = new ArrayList<String>();
		Predicate<String> rowPredicate = new RowPredicate(String.valueOf(row), new RowFunction(s -> Character.isDigit(s.charAt(0))));
		excelMap.keySet().stream().filter(rowPredicate).forEachOrdered(s -> rowList.add(excelMap.get(s)));
		return rowList;
	}

	@Override
	public List<String> getRow(int row, int skip) {
		List<String> rowList = new ArrayList<String>();
		Predicate<String> rowPredicate = new RowPredicate(String.valueOf(row), new RowFunction(s -> Character.isDigit(s.charAt(0))));
		excelMap.keySet().stream().filter(rowPredicate).skip(skip)
				.forEachOrdered(s -> rowList.add(excelMap.get(s)));
		return rowList;
	}

	@Override
	public List<String> getRow(int row, int skip, int limit) {
		List<String> rowList = new ArrayList<String>();
		Predicate<String> rowPredicate = new RowPredicate(String.valueOf(row), new RowFunction(s -> Character.isDigit(s.charAt(0))));
		excelMap.keySet().stream().filter(rowPredicate)
				.skip(skip).limit(limit)
				.forEachOrdered(s -> rowList.add(excelMap.get(s)));
		return rowList;
	}

	@Override
	public long getTotalRowCells(int row) {
		return excelMap.keySet().stream().filter(new RowPredicate(String.valueOf(row), new RowFunction(s -> Character.isDigit(s.charAt(0))))).count();
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
	public void removeColumn(String column) {
		column.toUpperCase();
		StringBuilder newKey = new StringBuilder();
		excelMap.keySet().stream()
			.filter(key -> key.substring(0, column.length()).equals(column)).sorted()
			.forEach(key -> this.excelMap.remove(key));
		Map<String, String> tempMap = new LinkedHashMap<>();
		for (Iterator<Map.Entry<String, String>> it = excelMap.entrySet().iterator(); it.hasNext();) {
			Map.Entry<String, String> entry = it.next();
			tempMap.put(updateColumnKey(newKey, entry.getKey()), entry.getValue());
			it.remove();
		}
		excelMap = new LinkedHashMap<String, String>(tempMap);
		tempMap.clear(); tempMap = null;
	}

	protected String updateColumnKey(StringBuilder newKey, String oldKey) {
		if (oldKey.contains("AA")) {
			return "AZ" + oldKey.charAt(2);
		} else if (oldKey.length() == 2 && oldKey.contains("A")) {
			return oldKey;
		}

		newKey.setLength(0);
		parseCellPosition(newKey, oldKey, ch -> Character.isLetter(ch));

		int newKeyLength = newKey.length();
		int oldKeyLength = oldKey.length();
		if (newKeyLength == 1) {
			return String.valueOf((char) (newKey.charAt(0) - 1)) + 
					oldKey.subSequence(1, oldKeyLength);
		} else {
			return String.valueOf((char) newKey.charAt(0)) + 
					String.valueOf((char) (newKey.charAt(1) - 1)) +
					oldKey.subSequence(2, oldKeyLength);
		}
	}

	@Override
	public void removeRow(int row) {
		StringBuilder newKey = new StringBuilder();
		Predicate<String> rowPredicate = new RowPredicate(String.valueOf(row), new RowFunction(ch -> Character.isDigit(ch.charAt(0))));
		excelMap.keySet().stream().filter(rowPredicate).sorted().forEach(key -> excelMap.remove(key));
		
		Map<String, String> tempMap = new LinkedHashMap<>();
		for (Iterator<Map.Entry<String, String>> it = excelMap.entrySet().iterator(); it.hasNext();) {
			Map.Entry<String, String> entry = it.next();
			String key = entry.getKey();
			if (key.length() == 2 && Integer.parseInt(key.substring(1)) > row) {
				tempMap.put(updateRowKey(newKey, entry.getKey()), entry.getValue());
				it.remove();
			}
		}
		excelMap.putAll(tempMap);
		tempMap.clear(); tempMap = null;
	}

	protected String updateRowKey(StringBuilder newKey, String oldKey) {
		newKey.setLength(0);
		parseCellPosition(newKey, oldKey, ch -> Character.isDigit(ch));

		int updatedRow = Integer.parseInt(newKey.toString()) - 1;
		newKey.setLength(0);
		int oldKeyLength = oldKey.length();
		if (oldKeyLength == 2) {
			return newKey.append(oldKey.substring(0, 1)).append(updatedRow).toString();
		} else {
			return newKey.append(oldKey.substring(0, 2)).append(updatedRow).toString();
		}
	}

	protected StringBuilder parseCellPosition(StringBuilder newKey, String oldKey, Predicate<Character> predicate) {
		for (int i = 0; i < oldKey.length(); i++) {
			char ch = oldKey.charAt(i);
			if (predicate.test(ch)) {
				newKey.append(ch);
			}
		}
		return newKey;
	}

	@Override
	public void removeCell(String cell, Move move) {
		if (Move.UP.equals(move)) {
			
		}
		throw new UnsupportedOperationException("removeCell is not yet implemented.");
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
	public String toString() {
		StringBuilder builder = new StringBuilder();
		builder.append("Total Cells: ").append(getTotalCells());
		return builder.toString();
	}

	private class RowFunction implements ExcelFunction<String, String> {

		private StringBuilder builder = new StringBuilder();
		private Predicate<String> predicate;
		
		public RowFunction(Predicate<String> predicate) {
			this.predicate = predicate;
		}

		@Override
		public String function(String s) {
			builder.setLength(0);
			for (int i = 0; i < s.length(); i++) {
				char character = s.charAt(i);
				if (predicate.test(String.valueOf(character))) {
					builder.append(character);
				}
			}
			return builder.toString();
		}
		
	}

	private class RowPredicate implements Predicate<String> {

		private String row;
		private ExcelFunction<String, String> excelFunction;
	
		public RowPredicate(String row, ExcelFunction<String, String> excelFunction) {
			this.row = row;
			this.excelFunction = excelFunction;
		}

		@Override
		public boolean test(String s) {
			return row.equals(excelFunction.function(s));
		}
		
	}

}
