package com.excelparser.parser;

import java.io.File;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import com.excelparser.matrixmap.ExcelMatrixMap;
import com.excelparser.matrixmap.MatrixMap;

public class ExcelParser {

	public MatrixMap processOneSheet(File file) throws Exception {
		OPCPackage pkg = OPCPackage.open(file);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();
		MatrixMap excelMatrixMap = new ExcelMatrixMap();
		XMLReader parser = fetchSheetParser(sst, excelMatrixMap);

		// To look up the Sheet Name / Sheet Order / rID,
		// you need to process the core Workbook stream.
		// Normally it's of the form rId# or rSheet#
		InputStream sheet = r.getSheet("rId1");
		InputSource sheetSource = new InputSource(sheet);
		parser.parse(sheetSource);
		sheet.close();
		return excelMatrixMap;
	}

	public MatrixMap processAllSheets(File file) throws Exception {
		OPCPackage pkg = OPCPackage.open(file);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();
		MatrixMap excelMatrixMap = new ExcelMatrixMap();
		XMLReader parser = fetchSheetParser(sst, excelMatrixMap);

		Iterator<InputStream> sheets = r.getSheetsData();
		while (sheets.hasNext()) {
			System.out.println("Processing new sheet:\n");
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
			System.out.println("");
		}
		return excelMatrixMap;
	}

	public XMLReader fetchSheetParser(SharedStringsTable sst,
			MatrixMap excelMatrixMap) throws SAXException {
		XMLReader parser = XMLReaderFactory
				.createXMLReader("org.apache.xerces.parsers.SAXParser");
		ContentHandler handler = new SheetHandler(sst, excelMatrixMap);
		parser.setContentHandler(handler);
		return parser;
	}

	/**
	 * See org.xml.sax.helpers.DefaultHandler javadocs
	 */
	private static class SheetHandler extends DefaultHandler {

		private SharedStringsTable sst;
		private String cellValue;
		private boolean nextIsString;
		private String cell;
		private MatrixMap excelMatrixMap;

		public SheetHandler(SharedStringsTable sst, MatrixMap excelMatrixMap) {
			this.sst = sst;
			this.excelMatrixMap = excelMatrixMap;
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
				cellValue = new XSSFRichTextString(sst.getEntryAt(idx))
						.toString();
				nextIsString = false;
			}

			// v => contents of a cell
			// Output after we've seen the string contents
			if (name.equals("v")) {
				excelMatrixMap.put(cell, cellValue);
			}
		}

		public void characters(char[] ch, int start, int length)
				throws SAXException {
			cellValue += new String(ch, start, length);
		}
	}

}
