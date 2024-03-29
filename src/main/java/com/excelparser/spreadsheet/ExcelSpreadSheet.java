package com.excelparser.spreadsheet;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Predicate;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class ExcelSpreadSheet implements SpreadSheet<String, String> {

	private Map<String, String> excelMap;
	private int startRowNum;
	private int endRowNum;
	private int rowCount;

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
		List<String> columnList = new ArrayList<String>();
		excelMap.keySet().stream().filter(s -> s.substring(0, column.length()).equals(column))
				.forEachOrdered(s -> columnList.add(excelMap.get(s)));
		return columnList;
	}

	@Override
	public List<String> getColumn(String column, int skip) {
		List<String> columnList = new ArrayList<String>();
		excelMap.keySet().stream().filter(s -> s.substring(0, column.length()).equals(column))
				.skip(skip)
				.forEachOrdered(s -> columnList.add(excelMap.get(s)));
		return columnList;
	}

	@Override
	public List<String> getColumn(String column, int skip, int limit) {
		List<String> columnList = new ArrayList<String>();
		excelMap.keySet().stream().filter(s -> s.substring(0, column.length()).equals(column))
				.skip(skip).limit(limit)
				.forEachOrdered(s -> columnList.add(excelMap.get(s)));
		return columnList;
	}

	@Override
	public long getTotalColumnCells(String column) {
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
		StringBuilder newKey = new StringBuilder();
		excelMap.keySet().stream()
			.filter(key -> key.substring(0, column.length()).equals(column)).sorted()
			.forEach(key -> this.excelMap.remove(key));
		Map<String, String> tempMap = new LinkedHashMap<>();
		for (Iterator<Map.Entry<String, String>> it = excelMap.entrySet().iterator(); it.hasNext();) {
			Map.Entry<String, String> entry = it.next();
			String key = entry.getKey();
			if (isColumnAfter(column, key)) {
				tempMap.put(updateColumnKey(newKey, key), entry.getValue());
				it.remove();
			}
		}
		excelMap.putAll(tempMap);
		tempMap.clear(); tempMap = null;
	}

	protected boolean isColumnAfter(String column, String key) {
		List<String> columnName = new ArrayList<String>(Arrays.asList(key.split("[0-9]")));
		columnName.removeIf(s -> s instanceof String && (s == null || "".equals(s)));
		return columnName.get(0).compareTo(column) >= 1 ? true : false;
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

	protected String updateColumnKey(StringBuilder newKey, String oldKey, int updateCol) {
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
			return String.valueOf((char) (newKey.charAt(0) + updateCol)) + 
					oldKey.subSequence(1, oldKeyLength);
		} else {
			return String.valueOf((char) newKey.charAt(0)) + 
					String.valueOf((char) (newKey.charAt(1) + updateCol)) +
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
			if (isRowAfter(row, key)) {
				tempMap.put(updateRowKey(newKey, key), entry.getValue());
				it.remove();
			}
		}
		excelMap.putAll(tempMap);
		tempMap.clear(); tempMap = null;
	}

	protected boolean isRowAfter(int row, String key) {
		List<String> rowNum = new ArrayList<String>(Arrays.asList(key.split("[a-zA-Z]")));
		rowNum.removeIf(s -> s instanceof String && (s == null || "".equals(s)));
		return Integer.parseInt(rowNum.get(0)) > row;
	}

	protected boolean compareKeys(String cellPos, String key, String regex) {
		List<String> capturedPosition = new ArrayList<String>(Arrays.asList(key.split(regex)));
		capturedPosition.removeIf(s -> s instanceof String && (s == null || "".equals(s)));
		int compare = capturedPosition.get(0).compareTo(cellPos);
		return compare >= 0 ? true : false;
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

	protected String updateRowKey(StringBuilder newKey, String oldKey, int updateRow) {
		newKey.setLength(0);
		parseCellPosition(newKey, oldKey, ch -> Character.isDigit(ch));

		int updatedRow = Integer.parseInt(newKey.toString()) + updateRow;
		newKey.setLength(0);
		int oldKeyLength = oldKey.length();
		if (oldKeyLength == 2) {
			return newKey.append(oldKey.substring(0, 1)).append(updatedRow).toString();
		} else {
			return newKey.append(oldKey.substring(0, 2)).append(updatedRow).toString();
		}
	}

	protected StringBuilder parseCellPosition(StringBuilder newKey, String oldKey, 
			Predicate<Character> predicate) {
		for (int i = 0; i < oldKey.length(); i++) {
			char ch = oldKey.charAt(i);
			if (predicate.test(ch)) {
				newKey.append(ch);
			}
		}
		return newKey;
	}

	@Override
	public void removeCell(String cell, Shift shift) {
		if (Shift.UP.equals(shift)) {
			removeCellShiftUp(cell);
		} else {
			removeCellShiftLeft(cell);
		}
	}

	private void removeCellShiftUp(String cell) {
		StringBuilder column = parseCellPosition(new StringBuilder(), cell, ch -> Character.isLetter(ch));
		String row = parseCellPosition(new StringBuilder(), cell, ch -> Character.isDigit(ch)).toString();
		int nextRow = Integer.parseInt(row);
		long totalRowsInColumn = getTotalColumnCells(column.toString());
		
		StringBuilder columnToUpdate = new StringBuilder();
		Map<String, String> tempMap = new LinkedHashMap<>();
		String keyWithNullNextVal = "";

		for (Iterator<Map.Entry<String, String>> it = excelMap.entrySet().iterator(); it.hasNext();) {
			Map.Entry<String, String> entry = it.next();
			String key = entry.getKey();
			columnToUpdate.setLength(0);

			parseCellPosition(columnToUpdate, key, ch -> Character.isLetter(ch));
			if (columnToUpdate.toString().equals(column.toString())) {
				if (compareKeys(row, key, "[a-zA-Z]")) {
					String value = excelMap.get(column.toString() + (++nextRow));
					if (value != null) {
						tempMap.put(keyWithNullNextVal != "" ? keyWithNullNextVal : key, value);
						keyWithNullNextVal = "";
					} else if ((nextRow - 1) == totalRowsInColumn) {
						columnToUpdate.setLength(0);
						tempMap.put(updateRowKey(columnToUpdate, key, -1), excelMap.get(key));
					} else {
						columnToUpdate.setLength(0);
						keyWithNullNextVal = updateRowKey(columnToUpdate, key, 1);
					}
				} else {
					tempMap.put(key, entry.getValue());
				}
			} else {
				tempMap.put(key, entry.getValue());
			}
			it.remove();
		}
		excelMap.putAll(tempMap);
		tempMap.clear(); tempMap = null;
	}

	private void removeCellShiftLeft(String cell) {
		StringBuilder column = parseCellPosition(new StringBuilder(), cell, ch -> Character.isLetter(ch));
		String row = parseCellPosition(new StringBuilder(), cell, ch -> Character.isDigit(ch)).toString();
		int rowAsInt = Integer.parseInt(row);
		long totalRowsInColumn = getTotalRowCells(rowAsInt);

		StringBuilder rowToUpdate = new StringBuilder();
		Map<String, String> tempMap = new LinkedHashMap<>();
		String keyWithNullNextVal = "";

		for (Iterator<Map.Entry<String, String>> it = excelMap.entrySet().iterator(); it.hasNext();) {
			Map.Entry<String, String> entry = it.next();
			String key = entry.getKey();
			rowToUpdate.setLength(0);

			parseCellPosition(rowToUpdate, key, ch -> Character.isDigit(ch));
			if (rowToUpdate.toString().equals(row)) {
				if (compareKeys(column.toString(), key, "[0-9]")) {
					rowToUpdate.setLength(0);
					String value = excelMap.get(updateColumnKey(rowToUpdate, key, 1));
					if (value != null) {
						tempMap.put(keyWithNullNextVal != "" ? keyWithNullNextVal : key, value);
						keyWithNullNextVal = "";
					} else if ((rowAsInt - 1) == totalRowsInColumn) {
						rowToUpdate.setLength(0);
						tempMap.put(updateRowKey(rowToUpdate, key, -1), excelMap.get(key));
					} else {
						rowToUpdate.setLength(0);
						keyWithNullNextVal = updateRowKey(rowToUpdate, key, 1);
					}
				} else {
					tempMap.put(key, entry.getValue());
				}
			} else {
				tempMap.put(key, entry.getValue());
			}
			it.remove();
		}
		excelMap.putAll(tempMap);
		tempMap.clear(); tempMap = null;
	}

	@Override
	public void clear() {
		excelMap.clear();
		rowCount = 0;
		startRowNum = 0;
		endRowNum = 0;
		log.info("Sheet cleared");
	}

	@Override
	public int getStartRowNum() {
		return startRowNum;
	}

	@Override
	public void setStartRowNum(int startRowNum) {
		this.startRowNum = startRowNum;
	}

	@Override
	public int getEndRowNum() {
		return endRowNum;
	}

	@Override
	public void setEndRowNum(int endRowNum) {
		this.endRowNum = endRowNum;
	}

	@Override
	public void incrementRowCount() {
		rowCount++;
	}

	@Override
	public int getRowCount() {
		return rowCount;
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
