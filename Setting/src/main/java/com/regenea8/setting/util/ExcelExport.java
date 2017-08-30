package com.regenea8.setting.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * Excel Export Class (.xlsx)
 * @version 1.0.0
 * @since 2017-05-19
 * @author jh.lee
 * 
 */
public class ExcelExport {
	
	/**
	 * Excel 생성 후 데이터 입력
	 * @param list
	 * @return
	 * @throws IOException
	 */
	public Workbook newExcel(List<LinkedHashMap<String, Object>> list) throws IOException {
		
		// 엑셀생성				
		Workbook wb = new SXSSFWorkbook(100);	// keep 100 rows in memory, exceeding rows will be flushed to disk
		String sheetname = "Sheet1";			// Excel Sheet 이름
		Sheet sh = wb.createSheet(sheetname);	// New Sheet 생성
	
		Iterator<LinkedHashMap<String, Object>> itLKHMap = list.iterator();
		Iterator<String> itStr = itLKHMap.next().keySet().iterator();
		
		// 컬럼명 입력
		Row row = sh.createRow(0);										// 행 생성												
		for(int cellNum = 0; itStr.hasNext(); cellNum++) {
			Cell cell = row.createCell(cellNum); 						// 열 생성
			String title = itStr.next();								// 데이터 꺼냄
			cell.setCellValue(title);	 								// 셀에 데이터를 입력
		}
		
		// 데이터 입력
		for(int rowNum = 1; itLKHMap.hasNext(); rowNum++) {
			row = sh.createRow(rowNum);									// 행 생성
			LinkedHashMap<String, Object> lkhmap = itLKHMap.next();		// List에서 Map을 꺼냄
			Iterator<String> key = lkhmap.keySet().iterator();			// Map에서 Key를 꺼냄
			for(int cellnum = 0; key.hasNext(); cellnum++) {
				Cell cell = row.createCell(cellnum); 					// 열 생성
				String cellValue = (String) lkhmap.get(key.next());		// 데이터 꺼냄
				cell.setCellValue(cellValue);	 						// 전달받은 데이터가 들어가는 부분
			}
		}
		
		return wb;
	
	}
	
	/**
	 * 생성한 Excel을 File로 저장
	 * @param list
	 * @param filePath
	 * @throws IOException
	 */
	public void excelSave(List<LinkedHashMap<String, Object>> list, String filePath) throws IOException {
		
		Workbook wb = newExcel(list);
		FileOutputStream out = new FileOutputStream(filePath);	// 경로와 파일 생성
		wb.write(out);											// 파일에 내용을 작성
		
		out.close();
		
	}
	
	/**
	 * 저장한 Excel을 Client에게 복사 및 다운로드
	 * @param list
	 * @param filePath
	 * @param newFileName
	 * @param response
	 * @throws IOException
	 */
	public void excelCopyResponse(List<LinkedHashMap<String, Object>> list, String filePath, String newFileName, HttpServletResponse response) throws IOException {
		
		excelSave(list, filePath);
		
		File file = new File(filePath);							// 불러올 파일명
		InputStream is = new FileInputStream(file);							

		String fileName = newFileName;										// 파일명
		fileName = new String(fileName.getBytes("UTF-8"), "ISO-8859-1");	// 파일명 한글 Encoding
		String fileExtension = ".xlsx";										// 확장자

		int bytelength = (int) file.length();
		byte fileByte[] = new byte[bytelength];
		
		response.reset(); // Some JSF component library or some Filter might have set some headers in the buffer beforehand. We want to get rid of them, else it may collide.
	    response.setContentType("application/octet-stream"); // Check http://www.iana.org/assignments/media-types for all types. Use if necessary ServletContext#getMimeType() for auto-detection based on filename.
	    response.setHeader("Content-Disposition", "attachment;filename=\"" + fileName + fileExtension + "\""); // The Save As popup magic is done here. You can give it any filename you want, this only won't work in MSIE, it will use current request URL as filename instead.

	    // Write file to response.
	    OutputStream output = response.getOutputStream();
	   
	    for (int nChunk = is.read(fileByte); nChunk!=-1; nChunk = is.read(fileByte)) {
	    	output.write(fileByte);
	    }
	    
	    is.close();
	    output.close();	
	    
	}
	
	/**
	 * 생성한 Excel을 File로 생성하지 않고 Client에게 다운로드 
	 * @param list
	 * @param newFileName
	 * @param response
	 * @throws IOException
	 */
	public void excelResponse(List<LinkedHashMap<String, Object>> list, String newFileName, HttpServletResponse response) throws IOException {
		
		Workbook wb = newExcel(list);
				
		String fileName = newFileName;										// 파일명
		fileName = new String(fileName.getBytes("UTF-8"), "ISO-8859-1");	// 파일명 한글 Encoding
		String fileExtension = ".xlsx";										// 확장자
	
		response.reset(); // Some JSF component library or some Filter might have set some headers in the buffer beforehand. We want to get rid of them, else it may collide.
	    response.setContentType("application/octet-stream"); // Check http://www.iana.org/assignments/media-types for all types. Use if necessary ServletContext#getMimeType() for auto-detection based on filename.
	    response.setHeader("Content-Disposition", "attachment;filename=\"" + fileName + fileExtension + "\""); // The Save As popup magic is done here. You can give it any filename you want, this only won't work in MSIE, it will use current request URL as filename instead.

	    // Write file to response.
	    OutputStream output = response.getOutputStream();
		wb.write(output);
	    
	    output.close();		
	    
	}
	
}
