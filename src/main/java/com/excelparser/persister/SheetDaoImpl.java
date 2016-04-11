package com.excelparser.persister;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;

import javax.sql.DataSource;

import com.excelparser.matrixmap.Matrix;

public class SheetDaoImpl<K, V> implements SheetDao<K, V> {

	private DataSource dataSource;
	private String createSql;

	public SheetDaoImpl() {
		// TODO Auto-generated constructor stub
	}

	public SheetDaoImpl(final DataSource dataSource, final String createSql) {
		this.dataSource = dataSource;
		this.createSql = createSql;
	}

	@Override
	public int create(final Matrix<K, V> spreadasheet) {
		Connection conn = null;
		PreparedStatement pst = null;
		try {
			conn = dataSource.getConnection();
			if (conn != null) {
				pst = conn.prepareStatement(createSql);
				for(int i = 0; i < spreadasheet.getRowCount(); i++) {
					
				}
				spreadasheet.clear();
				return pst.executeUpdate();
			}
		} catch (SQLException e) {
			
		} finally {
			if (pst != null) {
				try {
					pst.close();
				} catch (SQLException e) {
					// TODO deal with this
				}
			}
			if (conn != null) {
				try {
					conn.close();
				} catch (SQLException e) {
					// TODO deal with this
				}
			}
		}
		return 0;
	}

}
