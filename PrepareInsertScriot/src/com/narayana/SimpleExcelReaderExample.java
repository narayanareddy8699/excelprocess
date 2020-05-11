package com.narayana;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SimpleExcelReaderExample {

	public static void main(String[] args) throws IOException {
		String excelFilePath = ".xlsx";
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		Workbook workbook = new XSSFWorkbook(inputStream);
		Iterator<Sheet> sheetIterator = workbook.iterator();
		FileWriter fileWriter =null;
		BufferedWriter writer = null;
		while (sheetIterator.hasNext()) {
			Sheet sheet=sheetIterator.next();
			Iterator<Row> iterator = sheet.iterator();
			DataFormatter dataFormatter = new DataFormatter();
			fileWriter=new FileWriter("output.txt");
			writer=new BufferedWriter(fileWriter);
			int i=0;
			while(i<4) {
				iterator.next();
				i++;
			}			
			while (iterator.hasNext()) {
				Row nextRow = iterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();
				StringBuilder Insert = new StringBuilder(
						"UPDATE TB_CJ_INDUSTRYTYPES SET HIGH_RISK_INDUSTRY= ");
					Cell cell = cellIterator.next();
					String cellValue = dataFormatter.formatCellValue(cell);
					Insert.append("Greenlane".equalsIgnoreCase(cellValue.trim())?0:1 +"WHERE DBSIC_CODE=");
					Cell cell2 = cellIterator.next();
					String cellValue2 = dataFormatter.formatCellValue(cell);
					Insert.append(cellValue2.trim()+";");
				writer.write(Insert.toString() + "\n");
			}
			writer.close();
		}
		workbook.close();
		inputStream.close();

	}

}