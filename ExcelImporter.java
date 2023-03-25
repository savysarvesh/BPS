package com.marksheet.bps;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelImporter {
	public static void main(String args[]) throws IOException {
	
		 readData(args);
	}

	public static void readData(String args[]) throws IOException {

		String xlFileInput = args[0];
		String templateFile = args[1];
		String targetLocation = args[2];
		Map<Integer, List<String>> mapData = new HashMap<>();
		// obtaining input bytes from a file
		try {
			FileInputStream fis = new FileInputStream(new File(xlFileInput));
			// creating Workbook instance that refers to .xlsx file
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheetAt(0); // creating a Sheet object to retrieve object
			Iterator<Row> itr = sheet.iterator(); // iterating over excel file
			int count = 0;
			while (itr.hasNext()) {
				// System.out.println("ROW" + count);
				Row row = itr.next();

				Iterator<Cell> cellIterator = row.cellIterator(); // iterating over each column
				while (cellIterator.hasNext()) {
					String cellValue = "";
					Cell cell = cellIterator.next();

					switch (cell.getCellType()) {
					case STRING: // field that represents string cell type
						cellValue = cell.getStringCellValue();
						break;
					case NUMERIC: // field that represents number cell type
						cellValue = String.valueOf(cell.getNumericCellValue());
						break;
					case FORMULA: // field that represents number cell type
						cellValue = String.valueOf(cell.getNumericCellValue());
						break;
					default:
					}
//					if (row.getRowNum() == 0 && cell.getColumnIndex() > 0 && cell.getColumnIndex() < 273) {
//						// System.out.println("cell:" + cell.getColumnIndex() + "value : " + cellValue);
//						mapData.put(cell.getColumnIndex(), new ArrayList<String>());
//					} else
					if (row.getRowNum() > 2 && cell.getColumnIndex() > 0) {
//						mapData.put(cell.getColumnIndex(), mapData.get(cell.getColumnIndex()).addA(cellValue));
						mapData.computeIfAbsent(cell.getColumnIndex(), k -> new ArrayList<>()).add(cellValue);
					}
				}
				System.out.println("");
				count++;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("Operation successfully");

		for (int i = 0; i < mapData.size(); i++) {
			String fileName = "";
			Map<Integer, String> targetData = new HashMap<>();
			for (Entry<Integer, List<String>> entry : mapData.entrySet()) {
				String value = entry.getValue().get(i);
				if (entry.getKey() == 1) {
					fileName = value;
					fileName = fileName.replace(" ", "");
				}
				targetData.put(entry.getKey(), value);

			}
			PdfConverter.addDataToTheFile(templateFile, targetLocation + "\\" + fileName + ".docx", targetData);
		}

	}
}
