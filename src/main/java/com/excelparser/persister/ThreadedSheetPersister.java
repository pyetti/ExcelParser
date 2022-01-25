package com.excelparser.persister;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;

import lombok.extern.slf4j.Slf4j;

import com.excelparser.spreadsheet.SpreadSheet;

@Slf4j
public class ThreadedSheetPersister<K, V> implements SheetPersister<K, V> {

	private final Database<K, V> database;
	private List<Future<Integer>> results;
	private SheetPersistenceRunnerFactory<K, V> runnerFactory;
	private final ExecutorService executor = Executors.newFixedThreadPool(25);

	public ThreadedSheetPersister(final Database<K, V> database) {
		this.database = database;
		results = new ArrayList<Future<Integer>>();
		runnerFactory = new SheetPersistenceRunnerFactory<K, V>();
	}

	public ThreadedSheetPersister(final Database<K, V> database, 
			SheetPersistenceRunnerFactory<K, V> runnerFactory) {
		this.database = database;
		results = new ArrayList<Future<Integer>>();
		this.runnerFactory = runnerFactory;
	}

	@Override
	public void persist(final SpreadSheet<K, V> spreadSheet) {
		results.add(executor.submit(runnerFactory.getSheetPersistenceRunner(database, spreadSheet)));
	}

	public void shutdownExecutorService() {
		log.info("Shutting down ExecutorService");
		executor.shutdown(); // Disable new tasks from being submitted
	   try {
	     // Wait a while for existing tasks to terminate
	     if (!executor.awaitTermination(60, TimeUnit.SECONDS)) {
	       executor.shutdownNow(); // Cancel currently executing tasks
	       // Wait a while for tasks to respond to being cancelled
	       if (!executor.awaitTermination(60, TimeUnit.SECONDS)) {
	           log.error("ExecutorService did not terminate");
	       } else {
	    	   log.info("ExecutorService is shutdown");
	       }
	     } else {
	    	 log.info("ExecutorService is shutdown");
	     }
	   } catch (InterruptedException ie) {
	     // (Re-)Cancel if current thread also interrupted
	     executor.shutdownNow();
	     // Preserve interrupt status
	     Thread.currentThread().interrupt();
	   }
	}

	@Override
	public int getRowsPersisted() {
		int rowsPersisted = 0;
		try {
			for (Future<Integer> future : results) {
				int result = future.get();
				rowsPersisted += result;
			}
			return rowsPersisted;
		} catch (ExecutionException | InterruptedException e) {
			log.error(e.getMessage(), e);
		}
		return rowsPersisted;
	}

}
