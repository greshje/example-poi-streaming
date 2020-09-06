package com.greshje.example.poi.streaming;

import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.monitorjbl.xlsx.StreamingReader;

public class PoiStreamingExample {

	private static final Logger log = LoggerFactory.getLogger(PoiStreamingExample.class);

	private static final String FILE_NAME = "/com/greshje/example/poi/streaming/test-file.xlsx";

	public static void main(String[] args) {
		log.info("Starting test...");
		log.info("Getting file");
		InputStream in = PoiStreamingExample.class.getResourceAsStream(FILE_NAME);
		log.info("Got file");
		StreamingReader reader = getReader(in, 0);
		log.info("File contents:");
		for (Row row : reader) {
			String rowString = "";
			for (Cell cell : row) {
				if (rowString != "") {
					rowString += ",";
				}
				// NEED A WAY TO GET A DATE WHERE APPROPRIATE HERE
				rowString += cell.getStringCellValue();
			}
			log.info(rowString);
		}
		log.info("Done.");
	}

	public static StreamingReader getReader(InputStream in, int sheetIndex) {
		try {
			StreamingReader reader = StreamingReader.builder()
					.rowCacheSize(100) // number of rows to keep in memory (defaults to 10)
					.bufferSize(4096) // buffer size to use when reading InputStream to file (defaults to 1024)
					.sheetIndex(sheetIndex) // index of sheet to use
					.read(in); // read the file
			return reader;
		} catch (Exception exp) {
			throw new RuntimeException(exp);
		}
	}

}
