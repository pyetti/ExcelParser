package com.excelparser.matrixmap;

import java.util.List;
import java.util.Map;
import java.util.function.Predicate;

import com.excelparser.spreadsheet.Shift;
import lombok.extern.slf4j.Slf4j;

public interface Matrix<K, V> {

	/**
	 * Returns a list of cell data filtered by the Predicate passed in.
	 * 
	 * @param predicate
	 * @return list of cell data
	 */
	List<V> getCells(Predicate<String> predicate);

	List<V> getColumn(K column);

	/**
	 * Returns a List of cell data in the specified column. The list will skip
	 * an amount of rows based on the passed in amount
	 * 
	 * @param Column
	 *            : the column to to return
	 * @param skip
	 *            : the number of rows in the column to skip. The returned will
	 *            stars at skip + 1
	 * @return List of cell data
	 */
	List<V> getColumn(K column, int skip);

	/**
	 * Returns a List of cell data in the specified column. The list will skip
	 * and return an amount of rows based on the passed in amounts
	 * 
	 * @param Column
	 *            : the column to to return
	 * @param skip
	 *            : the number of rows in the column to skip. The returned will
	 *            stars at skip + 1
	 * @param limit
	 *            : the number of total rows in the column to return after start
	 *            + 1
	 * @return List of cell data
	 */
	List<V> getColumn(K column, int skip, int limit);

	long getTotalColumnCells(K column);

	List<V> getRow(int row);

	/**
	 * Returns a List of cell data in the specified row. The list will skip an
	 * amount of columns based on the passed in amount
	 * 
	 * @param Column
	 *            : the column to to return
	 * @param skip
	 *            : the number of columns in the row to skip. The returned will
	 *            stars at skip + 1
	 * @return List of cell data
	 */
	List<V> getRow(int row, int skip);

	/**
	 * Returns a List of cell data in the specified rows. The list will skip and
	 * limit the return to amount of columns based on the passed in amounts
	 * 
	 * @param Column
	 *            : the column to to return
	 * @param skip
	 *            : the number of columns in the row to skip. The returned will
	 *            stars at skip + 1
	 * @param limit
	 *            : the number of total columns in the row to return after start
	 *            + 1
	 * @return List of cell data
	 */
	List<V> getRow(int row, int skip, int limit);

	long getTotalRowCells(int row);

	V getEntry(K cell);

	boolean clearEntry(K key);

	boolean isEmpty();

	/**
	 * Returns the total number of cells in the document that have data
	 */
	int getTotalCells();

	/**
	 * Tests if the given value exists in the cell.
	 */
	boolean containsEntry(K value);

	/**
	 * Tests if a particular cell position (e.g. A3) contains data
	 */
	boolean containsCell(K key);

	V put(K cell, V value);

	/**
	 * Return the first row of the spreadsheet with data
	 */
	List<V> getHeader();

	void clear();

	/**
	 * Removing columns always moves the columns to the left.
	 */
	void removeColumn(String column);

	/**
	 * Removing rows always moves rows up
	 */
	void removeRow(int row);

	/**
	 * removeCell has to update the sheet depending on the Shift enum passed in, 
	 * which will mimics what happens in Excel
	 * 
	 * @param key
	 * @param move
	 */
	void removeCell(K key, Shift move);

	Map<K, V> getSheet();

}
