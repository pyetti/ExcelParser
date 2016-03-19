package com.excelparser.parser;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.util.Arrays;

import org.junit.Test;

import com.excelparser.matrixmap.MatrixMap;

public class TestExcelParser {

	private static final File SMALL_EXCEL_FILE = new File("src/test/resources/small-excel.xlsx");
	private static final File LARGE_EXCEL_FILE = new File("src/test/resources/excel.xlsx");

	@Test
	public void testProcessOneSheetSmallExcelFile() throws Exception {
		String[] expectedSubColumn = {"colB row3", "colB row4"};

		long testStart = System.currentTimeMillis();
		ExcelParser excelParser = new ExcelParser();
		MatrixMap excelMatrixMap = excelParser.processOneSheet(SMALL_EXCEL_FILE);
		assertEquals("colB row5", excelMatrixMap.getCellValue("B5"));
		assertEquals(Arrays.asList(expectedSubColumn), excelMatrixMap.getColumn("B", 2, 2));
		assertEquals(5, excelMatrixMap.getColumnLength("B"));
		assertEquals("colC row3", excelMatrixMap.getRow(3).get(2));
		System.out.println("Total time for small excel test (5 rows x 5 columns): " 
				+ (System.currentTimeMillis() - testStart) / 1000.0 + " seconds\n");
	}

	@Test
	public void testProcessOneSheetLargeExcelFile() throws Exception {
		String[] expectedSubColumn = {"colD row11", "colD row12"};
		
		long testStart = System.currentTimeMillis();
		ExcelParser excelParser = new ExcelParser();
		long start = System.currentTimeMillis();
		MatrixMap excelMatrixMap = excelParser.processOneSheet(LARGE_EXCEL_FILE);
		System.out.println("Total time parsing spreadsheet: " 
							+ (System.currentTimeMillis() - start) / 1000.0 + " seconds");

		start = System.currentTimeMillis();
		assertEquals("colB row5043", excelMatrixMap.getCellValue("B5043"));
		assertEquals(Arrays.asList(expectedSubColumn), excelMatrixMap.getColumn("D", 10, 2));
		assertEquals(100000, excelMatrixMap.getColumnLength("B"));
		assertEquals("colF row2334", excelMatrixMap.getRow(2334).get(5));
		System.out.println("Total time retrieving all request data: "
							+ (System.currentTimeMillis() - start) / 1000.0 + " seconds");
		System.out.println("Total time for large excel test (100,000 rows x 10 columns): " 
							+ (System.currentTimeMillis() - testStart) / 1000.0 + " seconds\n");
	}

}
