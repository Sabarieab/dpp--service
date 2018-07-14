package com.dag.services.controller;


import java.io.IOException;


import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import com.dag.services.beans.ExcelToJson;
import com.google.gson.JsonObject;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

@Controller
public class ExcelController {

	/** 
	 * API to get EXCEL contents in a JSON format
	 * @param multipartFile
	 * @param rows
	 * @return String (Json)
	 * @throws IOException
	 * @throws EncryptedDocumentException
	 * @throws InvalidFormatException
	 * @author Muralidharan E, Siddharth G
	 */
	@RequestMapping(method = RequestMethod.POST, value="/getJsonFromExcel", headers=("content-type=multipart/*"))
	@ResponseBody
	String getJsonFromExcel(@RequestParam("file") MultipartFile multipartFile, @RequestParam("rows") Integer rows) throws IOException, EncryptedDocumentException, InvalidFormatException {

		if (multipartFile == null || multipartFile.isEmpty()) {
			return null;
		}
		JsonObject jsonObject = ExcelToJson.getExcelDataAsJsonObject(multipartFile.getInputStream(), rows);
		return jsonObject.toString();
	}	
}
