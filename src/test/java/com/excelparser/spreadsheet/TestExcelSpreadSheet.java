package com.excelparser.spreadsheet;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.util.List;

import org.junit.BeforeClass;
import org.junit.Test;

import com.excelparser.matrixmap.Matrix;

public class TestExcelSpreadSheet {

	private static Matrix<String, String> spreadSheet;

	@BeforeClass
	public static void setupTests() {
		spreadSheet = new ExcelSpreadSheet();
		spreadSheet.put("A1", "Jan1"); spreadSheet.put("B1", "Feb1"); spreadSheet.put("C1", "Mar1");
		spreadSheet.put("A2", "Jan2"); spreadSheet.put("B2", "Feb2"); spreadSheet.put("C2", "Mar2");
		spreadSheet.put("A3", "Jan3"); spreadSheet.put("B3", "Feb3"); spreadSheet.put("C3", "Mar3");
		spreadSheet.put("A4", "Jan4"); 								  spreadSheet.put("C4", "Mar4");
		spreadSheet.put("A5", "Jan5");
																	  spreadSheet.put("C6", "Jan6");
		spreadSheet.put("A7", "Jan7");
	}

	@Test
	public void testGetColumn() {
		List<String> columnA = spreadSheet.getColumn("A");
		assertEquals(6, columnA.size());
		assertEquals("Jan1", columnA.get(0));
	}

	@Test
	public void testGetColumnSkipRows() {
		List<String> columnA = spreadSheet.getColumn("A", 2);
		assertEquals(4, columnA.size());
		assertEquals("Jan3", columnA.get(0));
	}

	@Test
	public void testGetColumnSkipRowsLimitRows() {
		List<String> columnA = spreadSheet.getColumn("A", 2, 4);
		assertEquals(4, columnA.size());
		assertEquals("Jan3", columnA.get(0));
	}

	@Test
	public void testGetTotalColumnCells() {
		long totalCells = spreadSheet.getTotalColumnCells("C");
		assertEquals(5, totalCells);
	}

	@Test
	public void testGetRow() {
		List<String> row2 = spreadSheet.getRow(2);
		assertEquals(3, row2.size());
		assertEquals("Feb2", row2.get(1));
	}

	@Test
	public void testGetRowSkipColumns() {
		List<String> row2 = spreadSheet.getRow(2, 2);
		assertEquals(1, row2.size());
		assertEquals("Mar2", row2.get(0));
	}

	@Test
	public void testGetRowSkipColumnsLimitColumns() {
		List<String> row3 = spreadSheet.getRow(3, 1, 3);
		assertEquals(2, row3.size());
		assertEquals("Feb3", row3.get(0));
	}

	@Test
	public void testGetTotalRowCells() {
		assertEquals(2, spreadSheet.getTotalRowCells(4));
	}

	@Test
	public void testGetEntry() {
		assertEquals("Mar2", spreadSheet.getEntry("C2"));
	}

	@Test
	public void testClearEntry() {
		Matrix<String, String> spreadSheet = new ExcelSpreadSheet();
		spreadSheet.put("A1", "Jan1"); spreadSheet.put("B1", "Feb1");
		assertTrue(spreadSheet.clearEntry("B1"));
		assertEquals("", spreadSheet.getEntry("B1"));
		assertFalse(spreadSheet.clearEntry("D1"));
	}

	@Test
	public void testPut() {
		Matrix<String, String> spreadSheet = new ExcelSpreadSheet();
		spreadSheet.put("A1", "Jan1"); spreadSheet.put("B1", "Feb1");
		spreadSheet.put("C2", "Mar2");
	}

	@Test
	public void testGetTotalCells() {
		assertEquals(14, spreadSheet.getTotalCells());
	}

	@Test
	public void testIsEmpty() {
		Matrix<String, String> spreadSheet2 = new ExcelSpreadSheet();
		assertTrue(spreadSheet2.isEmpty());
		assertFalse(spreadSheet.isEmpty());
	}

	@Test
	public void testContainsEntry() {
		assertTrue(spreadSheet.containsEntry("Feb1"));
		assertFalse(spreadSheet.containsEntry("Dec25"));
	}

	@Test
	public void testContainsCell() {
		assertTrue(spreadSheet.containsCell("B1"));
		assertFalse(spreadSheet.containsCell("F10"));
	}

	@Test
	public void testGetHeader() {
		List<String> header = spreadSheet.getHeader();
		assertEquals(3, header.size());
		assertEquals("Jan1", header.get(0));
		assertEquals("Feb1", header.get(1));
		assertEquals("Mar1", header.get(2));
	}

	@Test
	public void testRemoveColumn() {
		Matrix<String, String> spreadSheet = new ExcelSpreadSheet();
		spreadSheet.put("A1", "Jan1"); spreadSheet.put("B1", "Feb1"); spreadSheet.put("C1", "Mar1"); spreadSheet.put("D1", "Mar1");
		spreadSheet.put("A2", "Jan2"); spreadSheet.put("B2", "Feb2"); spreadSheet.put("C2", "Mar2"); spreadSheet.put("D2", "Mar1");
		spreadSheet.put("A3", "Jan3"); spreadSheet.put("B3", "Feb3"); spreadSheet.put("C3", "Mar3"); spreadSheet.put("D3", "Mar1");
		spreadSheet.put("A4", "Jan4"); 								  spreadSheet.put("C4", "Mar4");
		spreadSheet.put("A5", "Jan5");
																	  spreadSheet.put("C6", "Mar6"); spreadSheet.put("D4", "Mar1");
		spreadSheet.put("A7", "Jan7");

		List<String> colA = spreadSheet.getColumn("A");
		List<String> colD = spreadSheet.getColumn("D", 0, 10);
		String d2 = spreadSheet.getEntry("D2");
		spreadSheet.removeColumn("C");
		assertEquals(colD, spreadSheet.getColumn("C", 0, 10));
		assertEquals(d2, spreadSheet.getEntry("C2"));
		assertEquals(colA, spreadSheet.getColumn("A"));
	}

	@Test
	public void testIsColumnLessThan() {
		ExcelSpreadSheet spreadSheet = new ExcelSpreadSheet();
		assertTrue(spreadSheet.isColumnAfter("D", "F10"));
		assertTrue(spreadSheet.isColumnAfter("AD", "AF120"));
		assertFalse(spreadSheet.isColumnAfter("D", "B2"));
		assertFalse(spreadSheet.isColumnAfter("AD", "AB120"));
	}

	@Test
	public void testRemoveRow() {
		Matrix<String, String> spreadSheet = new ExcelSpreadSheet();
		spreadSheet.put("A1", "Jan1"); spreadSheet.put("B1", "Feb1"); spreadSheet.put("C1", "Mar1");
		spreadSheet.put("A2", "Jan2"); spreadSheet.put("B2", "Feb2"); spreadSheet.put("C2", "Mar2");
		spreadSheet.put("A3", "Jan3"); spreadSheet.put("B3", "Feb3"); spreadSheet.put("C3", "Mar3");
		spreadSheet.put("A4", "Jan4"); 								  spreadSheet.put("C4", "Mar4");
		spreadSheet.put("A5", "Jan5");
																	  spreadSheet.put("C6", "Mar6");
		spreadSheet.put("A7", "Jan7");

		List<String> row4 = spreadSheet.getRow(4, 0, 10);
		spreadSheet.removeRow(3);
		assertEquals(row4, spreadSheet.getRow(3));
	}

	@Test
	public void testUpdateRowKey() {
		ExcelSpreadSheet spreadSheet = new ExcelSpreadSheet();
		assertEquals("A3", spreadSheet.updateRowKey(new StringBuilder(), "A4"));
		assertEquals("AB3", spreadSheet.updateRowKey(new StringBuilder(), "AB4"));
	}

	@Test
	public void testParseCellPositionForColumn() {
		ExcelSpreadSheet spreadSheet = new ExcelSpreadSheet();
		StringBuilder sb = spreadSheet.parseCellPosition(new StringBuilder(), "B2", (Character c) -> Character.isLetter(c));
		assertEquals("B", sb.toString());

		StringBuilder sb2 = spreadSheet.parseCellPosition(new StringBuilder(), "BC2", (Character c) -> Character.isLetter(c));
		assertEquals("BC", sb2.toString());
	}

	@Test
	public void testIsRowLessThan() {
		ExcelSpreadSheet spreadSheet = new ExcelSpreadSheet();
		assertTrue(spreadSheet.isRowAfter(5, "A10"));
		assertFalse(spreadSheet.isRowAfter(150, "AB120"));
	}

	@Test
	public void testParseCellPositionForRow() {
		ExcelSpreadSheet spreadSheet = new ExcelSpreadSheet();
		StringBuilder sb = spreadSheet.parseCellPosition(new StringBuilder(), "B2", (Character c) -> Character.isDigit(c));
		assertEquals("2", sb.toString());

		StringBuilder sb2 = spreadSheet.parseCellPosition(new StringBuilder(), "AB100", (Character c) -> Character.isDigit(c));
		assertEquals("100", sb2.toString());
	}

	@Test
	public void testRemoveCellMoveUp() {
		Matrix<String, String> spreadSheet = new ExcelSpreadSheet();
		spreadSheet.put("A1", "Jan1"); spreadSheet.put("B1", "Feb1"); spreadSheet.put("C1", "Mar1");
		spreadSheet.put("A2", "Jan2"); spreadSheet.put("B2", "Feb2"); spreadSheet.put("C2", "Mar2");
		spreadSheet.put("A3", "Jan3"); spreadSheet.put("B3", "Feb3"); spreadSheet.put("C3", "Mar3");
		spreadSheet.put("A4", "Jan4"); 								  spreadSheet.put("C4", "Mar4");
		spreadSheet.put("A5", "Jan5");
																	  spreadSheet.put("C6", "Mar6");
		spreadSheet.put("A7", "Jan7");
																	  spreadSheet.put("C8", "Mar8");

		spreadSheet.removeCell("C2", Shift.UP);
		assertEquals("Mar3", spreadSheet.getEntry("C2"));
		assertEquals("Mar4", spreadSheet.getEntry("C3"));
		assertEquals("Mar6", spreadSheet.getEntry("C5"));
	}

	@Test
	public void testRemoveCellMoveLeft() {
		Matrix<String, String> spreadSheet = new ExcelSpreadSheet();
		spreadSheet.put("A1", "Jan1"); spreadSheet.put("B1", "Feb1"); spreadSheet.put("C1", "Mar1");
		spreadSheet.put("A2", "Jan2"); spreadSheet.put("B2", "Feb2"); spreadSheet.put("C2", "Mar2");
		spreadSheet.put("A3", "Jan3"); spreadSheet.put("B3", "Feb3"); spreadSheet.put("C3", "Mar3");
		spreadSheet.put("A4", "Jan4"); 								  spreadSheet.put("C4", "Mar4");
		spreadSheet.put("A5", "Jan5");
																	  spreadSheet.put("C6", "Mar6");
		spreadSheet.put("A7", "Jan7");

		spreadSheet.removeCell("B2", Shift.LEFT);
		assertEquals("Mar2", spreadSheet.getEntry("B2"));
		assertEquals("Feb3", spreadSheet.getEntry("B3"));
	}

	@Test
	public void testCompareCells() {
		ExcelSpreadSheet spreadSheet = new ExcelSpreadSheet();
		spreadSheet.put("A1", "Jan1"); spreadSheet.put("B1", "Feb1"); spreadSheet.put("C1", "Mar1");
		spreadSheet.put("A2", "Jan2"); spreadSheet.put("B2", "Feb2"); spreadSheet.put("C2", "Mar2");
		
		assertTrue(spreadSheet.compareKeys("1", "B2", "[a-zA-Z]"));
		assertTrue(spreadSheet.compareKeys("A", "B2", "[0-9]"));
		assertFalse(spreadSheet.compareKeys("3", "B2", "[a-zA-Z]"));
		assertFalse(spreadSheet.compareKeys("C", "B2", "[a-zA-Z]"));
	}

	@Test
	public void testClear() {
		Matrix<String, String> spreadSheet = new ExcelSpreadSheet();
		spreadSheet.put("A1", "Jan1"); spreadSheet.put("B1", "Feb1"); spreadSheet.put("C1", "Mar1");
		spreadSheet.clear();
		assertTrue(spreadSheet.isEmpty());
		assertEquals(0, spreadSheet.getTotalCells());
	}

	@Test
	public void testToString() {
		assertEquals("Total Cells: 14", spreadSheet.toString());
	}

}
