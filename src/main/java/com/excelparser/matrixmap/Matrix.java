package com.excelparser.matrixmap;

import java.util.Collection;
import java.util.List;

import com.excelparser.spreadsheet.Move;

public interface Matrix<K, V> {

	List<V> getColumn(K column);

	List<V> getColumn(K column, int start, int end);

	List<V> getColumn(K column, int start);

	long getTotalColumnCells(K column);

	List<V> getRow(int row);

	List<V> getRow(int row, int start, int end);

	List<V> getRow(int row, int start);

	long getTotalRowCells(int row);

	V getEntry(K cell);

	boolean clearEntry(K key);

	boolean isEmpty();

	int getTotalCells();

	boolean containsEntry(K value);

	boolean containsCell(K key);

	V put(K cell, V value);

	Collection<V> getHeader();

	void clear();

	V remove(K key, Move move);

}
