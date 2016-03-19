package com.excelparser.parser;

import java.io.File;

import org.junit.Test;

import com.excelparser.matrixmap.MatrixMap;

public class TestExcelParser {

	private static final File EXCEL_FILE = new File("src/test/resources/excel.xlsx");

	@Test
	public void testProcessOneSheet() throws Exception {
		ExcelParser excelParser = new ExcelParser();
		long start = System.currentTimeMillis();
		MatrixMap excelMatrixMap = excelParser.processOneSheet(EXCEL_FILE);
		System.out.println("Time parsing: " + (System.currentTimeMillis() - start));
		System.out.println(excelMatrixMap.getCellValue("B5"));
		System.out.println("Column B SubCol: " + excelMatrixMap.getColumn("B", 10, 2));
		System.out.println("Column B length: " + excelMatrixMap.getColumnLength("B"));
		System.out.println("Row 3: " + excelMatrixMap.getRow(3).get(2));
	}

}
