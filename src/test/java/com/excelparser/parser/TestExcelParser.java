package com.excelparser.parser;

import java.io.File;

import javax.sql.DataSource;

import lombok.extern.slf4j.Slf4j;

import com.excelparser.persister.SheetDaoImpl;
import com.excelparser.spreadsheet.SpreadSheet;
import org.junit.*;

import static org.junit.Assert.assertEquals;

@Slf4j
public class TestExcelParser {

	private static final File SMALL_EXCEL_FILE = new File("src/test/resources/small-excel.xlsx");
	private static final File LARGE_EXCEL_FILE = new File("src/test/resources/excel.xlsx");
	private static final File ASYM_EXCEL_FILE = new File("src/test/resources/asym-spreadsheet.xlsx");
	private static final int mb = 1024*1024;
	private static final Runtime runtime = Runtime.getRuntime();

	@BeforeClass
	public static void preTests() {
		log.info("Used Memory Before All Tests: " +
					((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
	}

	@Before
	public void preTest() {
		log.info("Used Memory Before Test: " +
				((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
	}

	@After
	public void postTest() {
		log.info("Used Memory After Test: " +
				((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
	}

	@AfterClass
	public static void postTests() {
		log.info("Used Memory After All Tests: " +
				((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB");
	}

	@Test
	public void testProcessOneSheetSmallExcelFile() throws Exception {
		ExcelParser excelParser = new ExcelParser();
		long testStart = System.currentTimeMillis();
		SpreadSheet<String, String> spreadSheet = excelParser.processOneSheet(SMALL_EXCEL_FILE);
		assertEquals(5, spreadSheet.getRowCount());
		assertEquals(25, spreadSheet.getTotalCells());
		log.info("Total time for small excel test (5 rows x 5 columns): "
				+ (System.currentTimeMillis() - testStart) / 1000.0 + " seconds\n");
	}

	@Test
	public void testProcessOneSheetLargeExcelFile() throws Exception {
		long testStart = System.currentTimeMillis();
		ExcelParser excelParser = new ExcelParser();
		SpreadSheet<String, String> spreadSheet = excelParser.processOneSheet(LARGE_EXCEL_FILE);
		assertEquals(100000, spreadSheet.getRowCount());
		assertEquals(1000000, spreadSheet.getTotalCells());
		log.info("Total time for large excel test (100,000 rows x 10 columns): "
							+ (System.currentTimeMillis() - testStart) / 1000.0 + " seconds");
	}

	@Test
	public void testAsymSpreadsheet() throws Exception {
		ExcelParser excelParser = new ExcelParser();
		long testStart = System.currentTimeMillis();
		SpreadSheet<String, String> spreadSheet = excelParser.processOneSheet(ASYM_EXCEL_FILE);
		assertEquals(20, spreadSheet.getRowCount());
		assertEquals(72, spreadSheet.getTotalCells());
		log.info("Total time for asym spreadsheet test (20 rows x 7 columns): "
				+ (System.currentTimeMillis() - testStart) / 1000.0 + " seconds\n");
	}

	@Test
	public void testProcessOneSheetSmallExcelFilePersistsDataSingleThread() throws Exception {
		ExcelParser excelParser = new ExcelParser();
		int rowsUploaded = excelParser.persistSheetData(SMALL_EXCEL_FILE, new StubSheetDao(null, null), 5);
		assertEquals(4, rowsUploaded);
	}

	@Test
	public void testProcessOneSheetSmallExcelFilePersistsDataMultiThreaded() throws Exception {
		ExcelParser excelParser = new ExcelParser();
		int rowsUploaded = excelParser.persistSheetData(SMALL_EXCEL_FILE, new StubSheetDao(null, null), 5);
		assertEquals(4, rowsUploaded);
	}

	private class StubSheetDao extends SheetDaoImpl<String, String> {

		public StubSheetDao(DataSource dataSource, String createSql) {
			super(dataSource, createSql);
		}

		@Override
		public int create(final SpreadSheet<String, String> spreadSheet) {
			return spreadSheet.getRowCount();
		}
		
	}

}
