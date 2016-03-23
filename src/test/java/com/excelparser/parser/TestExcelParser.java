package com.excelparser.parser;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.util.Arrays;

import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

import com.excelparser.matrixmap.Matrix;

public class TestExcelParser {

	private static final File SMALL_EXCEL_FILE = new File("src/test/resources/small-excel.xlsx");
	private static final File LARGE_EXCEL_FILE = new File("src/test/resources/excel.xlsx");
	private static final File ASYM_EXCEL_FILE = new File("src/test/resources/asym-spreadsheet.xlsx");
	private static final int mb = 1024*1024;
	private static final Runtime runtime = Runtime.getRuntime();

	@BeforeClass
	public static void preTests() {
		System.out.println("Used Memory: " + ((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
	}

	@AfterClass
	public static void postTests() {
		System.out.println("Used Memory: " + ((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB");
	}

	@Test
	public void testProcessOneSheetSmallExcelFile() throws Exception {
		String[] expectedSubColumn = {"B3", "B4"};

		ExcelParser excelParser = new ExcelParser();
		long testStart = System.currentTimeMillis();
		Matrix<String, String> spreadSheet = excelParser.processOneSheet(SMALL_EXCEL_FILE);
		spreadSheet.getRow(3).stream().filter((String s) -> s.equalsIgnoreCase("B3")).forEach((String s) -> System.out.println(s));
		assertEquals("B5", spreadSheet.getEntry("B5"));
		assertEquals(Arrays.asList(expectedSubColumn), spreadSheet.getColumn("B", 2, 2));
		assertEquals(5, spreadSheet.getTotalColumnCells("B"));
		assertEquals("C3", spreadSheet.getRow(3).get(2));
		assertEquals(25, spreadSheet.getTotalCells());
		System.out.println("Total time for small excel test (5 rows x 5 columns): " 
				+ (System.currentTimeMillis() - testStart) / 1000.0 + " seconds\n");
	}

	@Test
	public void testProcessOneSheetLargeExcelFile() throws Exception {
		System.out.println("Used Memory Large Excel: " + ((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB");
		String[] expectedSubColumn = {"D11", "D12"};
		
		long testStart = System.currentTimeMillis();
		ExcelParser excelParser = new ExcelParser();
		long start = System.currentTimeMillis();
		Matrix<String, String> spreadSheet = excelParser.processOneSheet(LARGE_EXCEL_FILE);
		System.out.println("Total time parsing large spreadsheet: " 
							+ (System.currentTimeMillis() - start) / 1000.0 + " seconds");

		start = System.currentTimeMillis();
		assertEquals("B5043", spreadSheet.getEntry("B5043"));
		assertEquals(Arrays.asList(expectedSubColumn), spreadSheet.getColumn("D", 10, 2));
		assertEquals(100000, spreadSheet.getTotalColumnCells("B"));
		assertEquals("F2334", spreadSheet.getRow(2334).get(5));
		assertEquals(1000000, spreadSheet.getTotalCells());
		System.out.println("Total time retrieving all request data: "
							+ (System.currentTimeMillis() - start) / 1000.0 + " seconds");
		System.out.println("Total time for large excel test (100,000 rows x 10 columns): " 
							+ (System.currentTimeMillis() - testStart) / 1000.0 + " seconds");
		System.out.println("Used Memory Large Excel: " + ((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
	}

	@Test
	public void testAsymSpreadsheet() throws Exception {
		String[] expectedSubColumn = {"C3", "C5"};

		ExcelParser excelParser = new ExcelParser();
		long testStart = System.currentTimeMillis();
		Matrix<String, String> spreadSheet = excelParser.processOneSheet(ASYM_EXCEL_FILE);
		assertEquals("B5", spreadSheet.getEntry("B5"));
		assertEquals(Arrays.asList(expectedSubColumn), spreadSheet.getColumn("C", 2, 2));
		assertEquals(12, spreadSheet.getTotalColumnCells("C"));
		assertTrue(spreadSheet.clearEntry("C6"));
		assertEquals(5, spreadSheet.getTotalRowCells(6));
		assertEquals("C3", spreadSheet.getRow(3).get(2));
		assertEquals(72, spreadSheet.getTotalCells());
		System.out.println("Total time for asym spreadsheet test: " 
				+ (System.currentTimeMillis() - testStart) / 1000.0 + " seconds\n");
	}

}
