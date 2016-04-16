package com.excelparser.persister;

import static org.junit.Assert.assertEquals;

import javax.sql.DataSource;

import org.junit.Test;

import com.excelparser.spreadsheet.ExcelSpreadSheet;
import com.excelparser.spreadsheet.SpreadSheet;

public class TestThreadedSheetPersister {

	@Test
	public void testPersist() {
		SpreadSheet<String, String> spreadSheet = new ExcelSpreadSheet();
		spreadSheet.put("A1", "Jan1"); spreadSheet.put("B1", "Feb1"); spreadSheet.put("C1", "Mar1");
		spreadSheet.put("A2", "Jan2"); spreadSheet.put("B2", "Feb2"); spreadSheet.put("C2", "Mar2");
		spreadSheet.put("A3", "Jan3"); spreadSheet.put("B3", "Feb3"); spreadSheet.put("C3", "Mar3");
		spreadSheet.incrementRowCount();
		spreadSheet.incrementRowCount();

		StubDatabase db = new StubDatabase(null, "");
		ThreadedSheetPersister<String, String> sheetPersister = new ThreadedSheetPersister<String, String>(
				db, new StubSheetPersistenceRunnerFactory());
		sheetPersister.persist(spreadSheet);
		assertEquals(2, sheetPersister.getRowsPersisted());
		sheetPersister.shutdownExecutorService();
	}

	private class StubDatabase extends Database<String, String> {

		public StubDatabase(DataSource dataSource, String sql) {
			super(dataSource, sql);
		}

	}

	private class StubSheetPersistenceRunnerFactory extends
			SheetPersistenceRunnerFactory<String, String> {

		@Override
		public SheetPersistenceRunner<String, String> getSheetPersistenceRunner(
				final Database<String, String> database,
				final SpreadSheet<String, String> spreadSheet) {
			return new StubSheetPersistenceRunner(database, spreadSheet);
		}

	}

	private class StubSheetPersistenceRunner extends
			SheetPersistenceRunner<String, String> {

		private SpreadSheet<String, String> spreadSheet;

		public StubSheetPersistenceRunner(Database<String, String> database,
				SpreadSheet<String, String> spreadSheet) {
			super(database, spreadSheet);
			this.spreadSheet = spreadSheet;
		}

		@Override
		public Integer call() throws Exception {
			return spreadSheet.getRowCount();
		}

	}

}
