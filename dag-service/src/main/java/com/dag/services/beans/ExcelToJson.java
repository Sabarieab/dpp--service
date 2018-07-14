package com.dag.services.beans;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;

public class ExcelToJson {

	/** 
	 * API to get EXCEL contents in a JSON format.
	 * @param fileInputStream
	 * @param rows JsonObject
	 * @return JsonObject
	 * @author Muralidharan E, Siddharth G
	 */
	public static JsonObject getExcelDataAsJsonObject(InputStream fileInputStream, Integer rows) {

		final int MAX_PREVIEW_COUNT = rows+1;

		JsonObject sheetsJsonObject = new JsonObject();
		Workbook workbook = null;

		try {
			workbook = new XSSFWorkbook(fileInputStream);
		} 
		catch (Exception e) {
			e.printStackTrace();
			//TODO Log Implementation
		}

		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

			JsonArray sheetArray = new JsonArray();
			ArrayList<String> columnNames = new ArrayList<String>();
			Sheet sheet = workbook.getSheetAt(i);
			Iterator<Row> sheetIterator = sheet.iterator();
			int count=1;
			while (sheetIterator.hasNext() && count++<=MAX_PREVIEW_COUNT) {

				Row currentRow = sheetIterator.next();
				JsonObject jsonObject = new JsonObject();

				if (currentRow.getRowNum() != 0) {

					for (int j = 0; j < columnNames.size(); j++) {

						if (currentRow.getCell(j) != null) {
							if (currentRow.getCell(j).getCellTypeEnum() == CellType.STRING) {
								jsonObject.addProperty(columnNames.get(j), currentRow.getCell(j).getStringCellValue());
							} else if (currentRow.getCell(j).getCellTypeEnum() == CellType.NUMERIC) {
								jsonObject.addProperty(columnNames.get(j), currentRow.getCell(j).getNumericCellValue());
							} else if (currentRow.getCell(j).getCellTypeEnum() == CellType.BOOLEAN) {
								jsonObject.addProperty(columnNames.get(j), currentRow.getCell(j).getBooleanCellValue());
							} else if (currentRow.getCell(j).getCellTypeEnum() == CellType.BLANK) {
								jsonObject.addProperty(columnNames.get(j), "");
							}
						} else {
							jsonObject.addProperty(columnNames.get(j), "");
						}

					}

					sheetArray.add(jsonObject);

				} else {
					// store column names
					for (int k = 0; k < currentRow.getPhysicalNumberOfCells(); k++) { //TODO Need to re-check this logic incase of null value scenarions... Muralidharan E
						if(currentRow.getCell(k)!=null)
						columnNames.add(currentRow.getCell(k).getStringCellValue());
					}
				}

			}

			sheetsJsonObject.add(workbook.getSheetName(i), sheetArray);

		}
		return sheetsJsonObject;
	}

}
