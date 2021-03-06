package com.excelparser.persister;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.List;

import javax.sql.DataSource;

import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;

import com.excelparser.spreadsheet.SpreadSheet;

public class SheetDaoImpl<K, V> implements SheetDao<K, V> {

	private final static Logger logger = Logger.getLogger(SheetDaoImpl.class.getName());
	private DataSource dataSource;
	private String insertQuery;

	public SheetDaoImpl(final DataSource dataSource, final String insertQuery) {
		this.dataSource = dataSource;
		this.insertQuery = insertQuery;
	}

	public SheetDaoImpl(final Database<K, V> database) {
		this.dataSource = database.getDataSource();
		this.insertQuery = database.getSql();
	}

	@Override
	public int create(final SpreadSheet<K, V> spreadSheet) {
		logger.info("Persisting " + spreadSheet.getRowCount() + " rows.");
		Connection conn = null;
		PreparedStatement pst = null;
		try {
			conn = dataSource.getConnection();
			if (conn != null) {
				int parameters = StringUtils.countMatches(insertQuery, "?");
				pst = conn.prepareStatement(insertQuery);
				int numRowsUploaded = 0;
				for(int rowNum = spreadSheet.getStartRowNum(); rowNum <= spreadSheet.getEndRowNum(); rowNum++) {
					List<V> row = spreadSheet.getRow(rowNum);
					if (row != null && row.size() == parameters) {
						for (int j = 0; j < parameters; j++) {
							pst.setObject(j+1, row.get(j));
						}
						pst.addBatch();
					}
				}
				spreadSheet.clear();
				numRowsUploaded += pst.executeBatch().length;
				pst.clearBatch();
				logger.info("Number of rows persisted: " + numRowsUploaded);
				return numRowsUploaded;
			}
			logger.error("Failed to connect to database.");
		} catch (SQLException e) {
			logger.error(e.getMessage(), e);
			return -1;
		} finally {
			close(conn, pst);
		}
		return 0;
	}

	private void close(Connection conn, PreparedStatement pst) {
		if (pst != null) {
			try {
				pst.close();
			} catch (SQLException e) {
				logger.error(e.getMessage(), e);
			}
		}
		if (conn != null) {
			try {
				conn.close();
			} catch (SQLException e) {
				logger.error(e.getMessage(), e);
			}
		}
	}

}
