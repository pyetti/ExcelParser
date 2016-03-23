package com.excelparser.spreadsheet;

import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

public class ExcelConsumer<K, V> implements Consumer<K> {

	private final List<V> list;
	private final Map<K, V> excelMap;

	public ExcelConsumer(List<V> list, Map<K, V> excelMap) {
		this.list = list;
		this.excelMap = excelMap;
	}

	@Override
	public void accept(K k) {
		list.add(excelMap.get(k));
	}

}
