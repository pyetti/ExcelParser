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
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import com.excelparser.persister.Database;
import com.excelparser.persister.SheetDao;
import com.excelparser.persister.SheetPersistenceHandler;
import com.excelparser.persister.SheetPersister;
import com.excelparser.persister.SimpleSheetPersister;
import com.excelparser.persister.ThreadedSheetPersister;
import com.excelparser.spreadsheet.ExcelSpreadSheet;
import com.excelparser.spreadsheet.SpreadSheet;

public class ExcelParser {

	private final static Logger logger = Logger.getLogger(ExcelParser.class);
	private static final int mb = 1024*1024;
	private static final Runtime runtime = Runtime.getRuntime();

	public SpreadSheet<String, String> processOneSheet(File file) throws Exception {
		logger.info("Starting processOneSheet");
		logger.info("Used Memory Before processOneSheet: " + 
				((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
		OPCPackage pkg = OPCPackage.open(file);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();

		SpreadSheet<String, String> spreadSheet = new ExcelSpreadSheet();
		XMLReader parser = fetchSheetParser(new SheetHandler(sst));

		// To look up the Sheet Name / Sheet Order / rID,
		// you need to process the core Workbook stream.
		// Normally it's of the form rId# or rSheet#
		InputStream sheet = r.getSheet("rId1");
		InputSource sheetSource = new InputSource(sheet);
		parser.parse(sheetSource);
		sheet.close();
		logger.info("Finished processOneSheet");
		logger.info("Used Memory After processOneSheet: " + 
				((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
		return spreadSheet;
	}

	// TODO test this
	public List<SpreadSheet<String, String>> processAllSheets(File file) throws Exception {
		logger.info("Starting processAllSheets");
		logger.info("Used Memory Before processAllSheets: " + 
				((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
		OPCPackage pkg = OPCPackage.open(file);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();

		List<SpreadSheet<String, String>> sheetList = new ArrayList<SpreadSheet<String, String>>();
		Iterator<InputStream> sheets = r.getSheetsData();
		while (sheets.hasNext()) {
			SpreadSheet<String, String> spreadSheet = new ExcelSpreadSheet();
			XMLReader parser = fetchSheetParser(new SheetHandler(sst, spreadSheet));
			logger.info("Processing new sheet:\n");
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
			sheetList.add(spreadSheet);
		}
		logger.info("Finished processAllSheets");
		logger.info("Used Memory After processAllSheets: " + 
				((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
		return sheetList;
	}

	public int persistSheetData(File file, SheetDao<String, String> sheetDao, 
			int numColumnsExpected) throws Exception {
		logger.info("Starting persistSheetData");
		logger.info("Used Memory Before persistSheetData: " + 
				((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
		OPCPackage pkg = OPCPackage.open(file);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();

		Database<String, String> database = new Database<String, String>(sheetDao);
		SheetPersister<String, String> sheetPersister = new SimpleSheetPersister<String, String>(database);
		SheetHandler handler = new SheetPersistenceHandler(sst, sheetPersister, numColumnsExpected);
		XMLReader parser = fetchSheetParser(handler);

		InputStream sheet = r.getSheet("rId1");
		InputSource sheetSource = new InputSource(sheet);
		parser.parse(sheetSource);
		sheet.close();
		if (!handler.getSpreadSheet().isEmpty()) {
			sheetPersister.persist(handler.getSpreadSheet());
		}
		logger.info("Finished persistSheetData");
		logger.info("Used Memory After persistSheetData: " + 
				((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
		return sheetPersister.getRowsPersisted();
	}

	public int persistSheetDataThreaded(File file, DataSource dataSource, String insertQuery, 
			int numColumnsExpected) throws Exception {
		logger.info("Starting persistSheetDataThreaded");
		logger.info("Used Memory Before Persist: " + 
				((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
		OPCPackage pkg = OPCPackage.open(file);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();

		Database<String, String> database = new Database<String, String>(dataSource, insertQuery);
		ThreadedSheetPersister<String, String> sheetPersister = new ThreadedSheetPersister<String, String>(database);
		SheetHandler handler = new SheetPersistenceHandler(sst, sheetPersister, numColumnsExpected);
		XMLReader parser = fetchSheetParser(handler);

		InputStream sheet = r.getSheet("rId1");
		InputSource sheetSource = new InputSource(sheet);
		parser.parse(sheetSource);
		sheet.close();
		if (!handler.getSpreadSheet().isEmpty()) {
			sheetPersister.persist(handler.getSpreadSheet());
		}
		int rowsPersisted = sheetPersister.getRowsPersisted();
		sheetPersister.shutdownExecutorService();
		logger.info("Finished persistSheetDataThreaded");
		logger.info("Used Memory After Persist: " + 
				((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
		return rowsPersisted;
	}

	public int persistSheetData(File file, SheetPersister<String, String> sheetPersister, 
			int numColumnsExpected) throws Exception {
		logger.info("Starting persistSheetData");
		logger.info("Used Memory Before persistSheetData: " + 
				((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
		OPCPackage pkg = OPCPackage.open(file);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();

		SheetHandler handler = new SheetPersistenceHandler(sst, sheetPersister, numColumnsExpected);
		XMLReader parser = fetchSheetParser(handler);

		InputStream sheet = r.getSheet("rId1");
		InputSource sheetSource = new InputSource(sheet);
		parser.parse(sheetSource);
		sheet.close();
		if (!handler.getSpreadSheet().isEmpty()) {
			sheetPersister.persist(handler.getSpreadSheet());
		}
		if (sheetPersister instanceof ThreadedSheetPersister) {
			((ThreadedSheetPersister<String, String>) sheetPersister).shutdownExecutorService();
		}
		logger.info("Starting persistSheetData");
		logger.info("Used Memory Before persistSheetData: " + 
				((runtime.totalMemory() - runtime.freeMemory()) / mb) + "MB\n");
		return sheetPersister.getRowsPersisted();
	}

	private XMLReader fetchSheetParser(DefaultHandler sheetHandler) throws SAXException {
		XMLReader parser = XMLReaderFactory
				.createXMLReader("org.apache.xerces.parsers.SAXParser");
		parser.setContentHandler(sheetHandler);
		return parser;
	}

}
