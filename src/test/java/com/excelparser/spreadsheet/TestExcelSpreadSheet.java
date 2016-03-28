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
		ExcelSpreadSheet spreadSheet = new ExcelSpreadSheet();
		spreadSheet.put("A1", "Jan1"); spreadSheet.put("B1", "Feb1");
		assertTrue(spreadSheet.clearEntry("B1"));
		assertEquals("", spreadSheet.getEntry("B1"));
		assertFalse(spreadSheet.clearEntry("D1"));
	}

	@Test
	public void testPut() {
		ExcelSpreadSheet spreadSheet = new ExcelSpreadSheet();
		spreadSheet.put("A1", "Jan1"); spreadSheet.put("B1", "Feb1");
		spreadSheet.put("C2", "Mar2");
	}

	@Test
	public void testGetTotalCells() {
		assertEquals(14, spreadSheet.getTotalCells());
	}

	@Test
	public void testIsEmpty() {
		ExcelSpreadSheet spreadSheet2 = new ExcelSpreadSheet();
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
		ExcelSpreadSheet spreadSheet = new ExcelSpreadSheet();
		spreadSheet.put("A1", "Jan1"); spreadSheet.put("B1", "Feb1"); spreadSheet.put("C1", "Mar1");
		spreadSheet.put("A2", "Jan2"); spreadSheet.put("B2", "Feb2"); spreadSheet.put("C2", "Mar2");
		spreadSheet.put("A3", "Jan3"); spreadSheet.put("B3", "Feb3"); spreadSheet.put("C3", "Mar3");
		spreadSheet.put("A4", "Jan4"); 								  spreadSheet.put("C4", "Mar4");
		spreadSheet.put("A5", "Jan5");
																	  spreadSheet.put("C6", "Jan6");
		spreadSheet.put("A7", "Jan7");
		System.out.println(spreadSheet.getSheet());
		
		List<String> colB = spreadSheet.getColumn("C", 0, 10);
		System.out.println(spreadSheet.getEntry("C2"));
		System.out.println(colB);
		spreadSheet.removeColumn("B");
		colB = spreadSheet.getColumn("B", 0, 10);
		System.out.println(colB);
		System.out.println(spreadSheet.getEntry("B2"));
		
		System.out.println(spreadSheet.getSheet());
	}

	@Test
	public void testRemoveRow() {
		
	}

	@Test
	public void testRemoveCell() {
		
	}

	@Test
	public void testClear() {
		ExcelSpreadSheet spreadSheet = new ExcelSpreadSheet();
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
