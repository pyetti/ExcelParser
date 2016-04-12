package com.excelparser.parser;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.sql.DataSource;

import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import com.excelparser.matrixmap.Matrix;
import com.excelparser.persister.Database;
import com.excelparser.persister.SheetDao;
import com.excelparser.persister.SheetPersister;
import com.excelparser.persister.SimpleSheetPersister;
import com.excelparser.persister.ThreadedSheetPersister;
import com.excelparser.spreadsheet.ExcelSpreadSheet;

public class ExcelParser {

	private final static Logger logger = Logger.getLogger(ExcelParser.class);

	public Matrix<String, String> processOneSheet(File file) throws Exception {
		OPCPackage pkg = OPCPackage.open(file);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();

		Matrix<String, String> spreadSheet = new ExcelSpreadSheet();
		XMLReader parser = fetchSheetParser(new SheetHandler(sst, spreadSheet));

		// To look up the Sheet Name / Sheet Order / rID,
		// you need to process the core Workbook stream.
		// Normally it's of the form rId# or rSheet#
		InputStream sheet = r.getSheet("rId1");
		InputSource sheetSource = new InputSource(sheet);
		parser.parse(sheetSource);
		sheet.close();
		return spreadSheet;
	}

	// TODO test this
	public List<Matrix<String, String>> processAllSheets(File file) throws Exception {
		OPCPackage pkg = OPCPackage.open(file);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();

		List<Matrix<String, String>> sheetList = new ArrayList<Matrix<String, String>>();
		Iterator<InputStream> sheets = r.getSheetsData();
		while (sheets.hasNext()) {
			Matrix<String, String> spreadSheet = new ExcelSpreadSheet();
			XMLReader parser = fetchSheetParser(new SheetHandler(sst, spreadSheet));
			logger.info("Processing new sheet:\n");
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
			sheetList.add(spreadSheet);
		}
		return sheetList;
	}

	public int persistSheetData(File file, SheetDao<String, String> sheetDao, 
			int numColumnsExpected) throws Exception {
		OPCPackage pkg = OPCPackage.open(file);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();

		Matrix<String, String> spreadSheet = new ExcelSpreadSheet();
		Database<String, String> database = new Database<String, String>(sheetDao);
		SheetPersister<String, String> sheetPersister = new SimpleSheetPersister<String, String>(database);
		SheetHandler handler = new SheetHandler(sst, spreadSheet, sheetPersister, numColumnsExpected);
		XMLReader parser = fetchSheetParser(handler);

		InputStream sheet = r.getSheet("rId1");
		InputSource sheetSource = new InputSource(sheet);
		parser.parse(sheetSource);
		sheet.close();
		if (!spreadSheet.isEmpty()) {
			sheetDao.create(spreadSheet);
		}
		return handler.getRowsPersisted();
	}

	public int persistSheetData(File file, DataSource dataSource, String sql, 
			int numColumnsExpected) throws Exception {
		OPCPackage pkg = OPCPackage.open(file);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();

		Matrix<String, String> spreadSheet = new ExcelSpreadSheet();
		Database<String, String> database = new Database<String, String>(dataSource, sql);
		ThreadedSheetPersister<String, String> sheetPersister = new ThreadedSheetPersister<String, String>(database);
		SheetHandler handler = new SheetHandler(sst, spreadSheet, sheetPersister, numColumnsExpected);
		XMLReader parser = fetchSheetParser(handler);

		InputStream sheet = r.getSheet("rId1");
		InputSource sheetSource = new InputSource(sheet);
		parser.parse(sheetSource);
		sheet.close();
		if (!spreadSheet.isEmpty()) {
			sheetPersister.persist(spreadSheet);
		}
		sheetPersister.shutdownExecutorService();
		return handler.getRowsPersisted();
	}

	public int persistSheetData(File file, SheetPersister<String, String> sheetPersister, 
			int numColumnsExpected) throws Exception {
		OPCPackage pkg = OPCPackage.open(file);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();

		Matrix<String, String> spreadSheet = new ExcelSpreadSheet();
		SheetHandler handler = new SheetHandler(sst, spreadSheet, sheetPersister, numColumnsExpected);
		XMLReader parser = fetchSheetParser(handler);

		InputStream sheet = r.getSheet("rId1");
		InputSource sheetSource = new InputSource(sheet);
		parser.parse(sheetSource);
		sheet.close();
		if (!spreadSheet.isEmpty()) {
			sheetPersister.persist(spreadSheet);
		}
		if (sheetPersister instanceof ThreadedSheetPersister) {
			((ThreadedSheetPersister<String, String>) sheetPersister).shutdownExecutorService();
		}
		return handler.getRowsPersisted();
	}

	private XMLReader fetchSheetParser(DefaultHandler sheetHandler) throws SAXException {
		XMLReader parser = XMLReaderFactory
				.createXMLReader("org.apache.xerces.parsers.SAXParser");
		parser.setContentHandler(sheetHandler);
		return parser;
	}

	/**
	 * See org.xml.sax.helpers.DefaultHandler javadocs
	 */
	private class SheetHandler extends DefaultHandler {

		private SharedStringsTable sst;
		private String cellValue;
		private boolean nextIsString;
		private String cell;
		private int row;
		private final Matrix<String, String> spreadSheet;
		private SheetPersister<String, String> sheetPersister;
		private boolean persist;
		private int maxRowsPerRead = 2;
		private int numColumnsExpected;
		private int numColumnsRead;
		private int rowsPersisted;
		private XSSFRichTextStringParser xssfRichTextString;

		public SheetHandler(SharedStringsTable sst, Matrix<String, String> spreadSheet, 
				SheetPersister<String, String> sheetPersister, int numColumnsExpected) {
			this.sst = sst;
			this.spreadSheet = spreadSheet;
			this.sheetPersister = sheetPersister;
			this.numColumnsExpected = numColumnsExpected;
			this.persist = true;
		}

		public SheetHandler(SharedStringsTable sst, Matrix<String, String> spreadSheet) {
			this.sst = sst;
			this.spreadSheet = spreadSheet;
			this.persist = false;
		}

		public void startElement(String uri, String localName, String name,
				Attributes attributes) throws SAXException {
			// c => cell
			if (name.equals("c")) {
				// Print the cell reference
				cell = attributes.getValue("r");
				// System.out.print(attributes.getValue("r") + " - ");
				// Figure out if the value is an index in the SST
				String cellType = attributes.getValue("t");
				if (cellType != null && cellType.equals("s")) {
					nextIsString = true;
				} else {
					nextIsString = false;
				}
			} else if (name.equals("row")) {
				row++;
				spreadSheet.incrementRowCount();
				if (persist) {
					numColumnsRead = 0;
				}
			}
			// Clear contents cache
			cellValue = "";
		}

		public void endElement(String uri, String localName, String name)
				throws SAXException {
			// Process the last contents as required.
			// Do now, as characters() may be called more than once
			if (nextIsString) {
				int idx = Integer.parseInt(cellValue);
				cellValue = xssfRichTextString.getString(sst.getEntryAt(idx));
				nextIsString = false;
			}

			// v => contents of a cell
			// Output after we've seen the string contents
			if (name.equals("v")) {
				if (persist) {
					doPersist();
				} else {
					spreadSheet.put(cell, cellValue);
				}
			}
		}

		private void doPersist() {
			if (row > 1) {
				spreadSheet.put(cell, cellValue);
				numColumnsRead++;
				if (row % maxRowsPerRead == 0 && numColumnsRead % numColumnsExpected == 0) {
					sheetPersister.persist(spreadSheet);
					rowsPersisted += sheetPersister.getRowsPersisted();
					spreadSheet.clear();
					numColumnsRead = 0;
				}
			}
		}

		public void characters(char[] ch, int start, int length)
				throws SAXException {
			cellValue += new String(ch, start, length);
		}

		public int getRowsPersisted() {
			return rowsPersisted;
		}

	}

}
