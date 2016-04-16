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

	private final static Logger logger = Logger.getLogger(SheetDaoImpl.class);
	private DataSource dataSource;
	private String createSql;

	public SheetDaoImpl(final DataSource dataSource, final String createSql) {
		this.dataSource = dataSource;
		this.createSql = createSql;
	}

	@Override
	public int create(final SpreadSheet<K, V> spreadSheet) {
		Connection conn = null;
		PreparedStatement pst = null;
		try {
			conn = dataSource.getConnection();
			if (conn != null) {
				int parameters = StringUtils.countMatches(createSql, "?");
				pst = conn.prepareStatement(createSql);
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
				return pst.executeBatch().length;
			}
		} catch (SQLException e) {
			logger.error(e.getMessage(), e);
		} finally {
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
		return 0;
	}

}
