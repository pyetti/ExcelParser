package com.excelparser.persister;

import javax.sql.DataSource;

// TODO There has to be a better name for this class
public class Database<K, V> {

	private DataSource dataSource;
	private String sql;
	private SheetDao<K, V> sheetDao;

	public Database(DataSource dataSource, String sql) {
		this.dataSource = dataSource;
		this.sql = sql;
	}

	public Database(SheetDao<K, V> sheetDao) {
		this.sheetDao = sheetDao;
	}

	public DataSource getDataSource() {
		return dataSource;
	}

	public String getSql() {
		return sql;
	}

	public SheetDao<K, V> getSheetDao() {
		return sheetDao;
	}

}
