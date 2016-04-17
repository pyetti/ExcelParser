package com.excelparser.parser;

/**
 * See org.xml.sax.helpers.DefaultHandler javadocs
 */
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import com.excelparser.spreadsheet.ExcelSpreadSheet;
import com.excelparser.spreadsheet.SpreadSheet;

public class SheetHandler extends DefaultHandler {

	private SharedStringsTable sst;
	private String cellValue;
	private boolean nextIsString;
	private String cell;
	private int row;
	private SpreadSheet<String, String> spreadSheet;
	private XSSFRichTextStringParser xssfRichTextString;
	private boolean nextIsDate;

	public SheetHandler(SharedStringsTable sst) {
		this.sst = sst;
		this.spreadSheet = new ExcelSpreadSheet();
		xssfRichTextString = new XSSFRichTextStringParser();
	}

	public SheetHandler(SharedStringsTable sst, SpreadSheet<String, String> spreadSheet) {
		this.sst = sst;
		this.spreadSheet = spreadSheet;
		xssfRichTextString = new XSSFRichTextStringParser();
	}

	public void startElement(String uri, String localName, String name,
			Attributes attributes) throws SAXException {
		// c => cell
		if (name.equals("c")) {
			nextIsDate = attributes.getValue("s") != null ? true : false;
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
			handleNewRow();
		}
		// Clear contents cache
		cellValue = "";
	}

	protected void handleNewRow() {
		row++;
		spreadSheet.incrementRowCount();
		updateSheetRowData();
	}

	private void updateSheetRowData() {
		if (spreadSheet.getStartRowNum() == 0) {
			spreadSheet.setStartRowNum(row);
		}
		spreadSheet.setEndRowNum(row);
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
			cellValue = nextIsDate ? getFormattedDateString() : cellValue;
			handleCellData();
		}
	}

	protected void handleCellData() {
		spreadSheet.put(cell, cellValue);
	}

	private String getFormattedDateString() {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		return sdf.format(DateUtil.getJavaDate(Double.parseDouble(cellValue)));
	}

	public void characters(char[] ch, int start, int length)
			throws SAXException {
		cellValue += new String(ch, start, length);
	}

	public SpreadSheet<String, String> getSpreadSheet() {
		return spreadSheet;
	}

	public String getCellValue() {
		return cellValue;
	}

	public String getCell() {
		return cell;
	}

}
