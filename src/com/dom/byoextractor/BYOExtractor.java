package com.dom.byoextractor;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BYOExtractor extends BYOExtExtractor{

	private static final String OUTPUT_FILE = "output/"+BYOExtractorUtils.fileName;
	private static XSSFWorkbook wb = new XSSFWorkbook();

	public static void main(String args[]) {
		BYOExtExtractor  exterior = new BYOExtExtractor();
		exterior.BYOExtExterior(wb);
		BYOIntExtractor interior = new BYOIntExtractor();
		interior.BYOIntExterior(wb);
		BYOEntExtractor entertainment = new BYOEntExtractor();
		entertainment.BYOEntExterior(wb);
		BYOPerfExtractor performance = new BYOPerfExtractor();
		performance.BYOPerfExterior(wb);
		BYOServExtractor service = new BYOServExtractor();
		service.BYOServExterior(wb);
		/**/
		try {
			FileOutputStream outputStream = new FileOutputStream(OUTPUT_FILE);
			wb.write(outputStream);
			outputStream.close();
			System.out.println("BYO sheet Extracted successfully");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}

