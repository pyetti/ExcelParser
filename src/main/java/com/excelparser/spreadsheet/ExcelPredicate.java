package com.excelparser.spreadsheet;

import java.util.function.Predicate;

public class ExcelPredicate<T, F, R> implements Predicate<F> {

	private final T t;
	private final ExcelFunction<F, R> excelFunction;

	public ExcelPredicate(T t, ExcelFunction<F, R> function) {
		this.t = t;
		this.excelFunction = function;
	}

	@Override
	public boolean test(F f) {
		return t.equals(excelFunction.function(f));
	}

}
