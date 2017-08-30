package com.regenea8.setting.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.multipart.MultipartFile;

/**
 * ExcelUtil Class
 * @version 0.9.0
 * @since 2017-07-05
 * @author JunHwan-Lee
 */
public class ExcelUtil {

	private static final Logger logger = LoggerFactory.getLogger(ExcelUtil.class);
	
	/**
	 * 
	 * @param originalName
	 * @param fileData
	 * @param uploadPath
	 * @return savedName
	 * @throws Exception
	 */
	public List<Map<String, Object>> main(MultipartFile multipartFile, String uploadPath, String[] columnNames) throws Exception {
		
		logger.info("originalName : " + multipartFile.getOriginalFilename());	// 파일명.확장자
		logger.info("size         : " + multipartFile.getSize());				// 파일 용량(byte)
		logger.info("contentType  : " + multipartFile.getContentType());		// 파일 종류
		
		String filePath = uploadfile(multipartFile.getOriginalFilename(), multipartFile.getBytes(), uploadPath);
		logger.info("filePath     : " + filePath);								// 파일 종류
		
		List<Map<String, Object>> list = excelImport(columnNames, filePath);
		
		return list;
	}
	
	/**
	 * 
	 * @param originalName
	 * @param fileData
	 * @param uploadPath
	 * @return savedName
	 * @throws Exception
	 */
	public String uploadfile(String originalName, byte[] fileData, String uploadPath) throws Exception {

		UUID uid = UUID.randomUUID();
		String savedName = uid.toString() + "_" + originalName;

		File target = new File(uploadPath, savedName);
		FileCopyUtils.copy(fileData, target);

		String filePath = uploadPath + savedName;
		
		return filePath;
	}

	/**
	 * Excel 가져오기
	 * @param arrColumns
	 * @param filePath
	 * @return list
	 * @throws Exception
	 */
	public List<Map<String, Object>> excelImport(String[] columnNames, String filePath) throws Exception {

		List<Map<String, Object>> list = null;
		Map<String, Object> map = null;

		Workbook workbook = null;
		FileInputStream fis = null;

		String fileExtension = filePath.substring(filePath.lastIndexOf(".")).toLowerCase();	// 확장자 추출

		try {

			if (".xls".equals(fileExtension)) {
				fis = new FileInputStream(filePath);
				workbook = new HSSFWorkbook(fis);
			} else if (".xlsx".equals(fileExtension)) { // 확장자 검사
				fis = new FileInputStream(filePath);
				workbook = new XSSFWorkbook(fis);
			} else {
				throw new Exception("지원하지 않는 확장자"); // My Exception으로 변경 예정
			}

			list = new ArrayList<Map<String,Object>>();

			int rowIndex = 0;
			int columnIndex = 0;

			Sheet sheet = workbook.getSheetAt(0);								// Sheet 순서
			int rows = sheet.getPhysicalNumberOfRows();							// Row 갯수

			for(rowIndex=0; rowIndex < rows; rowIndex++) {

				Row row = sheet.getRow(rowIndex);								// Row를 불러옴.

				if (row == null) {
					throw new Exception("Row 가져오기 실패.");
				}

				map = new HashMap<String, Object>();
				int cells = row.getPhysicalNumberOfCells();						// Cell 갯수

				for(columnIndex = 0; columnIndex < cells; columnIndex++) {       	
					Cell cell = row.getCell(columnIndex);						// Cell를 불러옴.
					String value = "";		

					if (cell.getCellType() == XSSFCell.CELL_TYPE_FORMULA) {
						value = cell.getCellFormula();
					} else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
						if (DateUtil.isCellDateFormatted(cell)) {
							SimpleDateFormat objSimpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
							value = "" + objSimpleDateFormat.format(cell.getDateCellValue());
						} else {
							value = "" + String.format("%.0f", new Double(cell.getNumericCellValue()));
						}
					} else if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
						value = cell.getStringCellValue() + "";
					} else if (cell.getCellType() == XSSFCell.CELL_TYPE_BLANK) {
						value = cell.getBooleanCellValue() + "";
					} else if (cell.getCellType() == XSSFCell.CELL_TYPE_ERROR) {
						value = cell.getErrorCellValue() + "";
					} else {
						throw new Exception("Cell 가져오기 실패.");
					}
					map.put(columnNames[columnIndex], value);
				}
				list.add(map);
			}

		} catch (NullPointerException npe) { 
			throw npe;
		} catch (Exception e) {
			throw e;
		} finally {
			try {
				if (fis != null) {
					fis.close();
				}
			} catch (Exception e) {
				throw e;
			}
		}

		return list;
	}

	/**
	 * Excel 생성 후 데이터 입력
	 * @param list
	 * @return workbook
	 * @throws IOException
	 */
	public Workbook createExcel(List<LinkedHashMap<String, Object>> list) throws IOException {

		// 엑셀생성				
		Workbook workbook = new SXSSFWorkbook(100);	// keep 100 rows in memory, exceeding rows will be flushed to disk
		String sheetname = "Sheet1";			// Excel Sheet 이름
		Sheet sh = workbook.createSheet(sheetname);	// New Sheet 생성

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

		return workbook;

	}

	/**
	 * 생성한 Excel을 File로 저장
	 * @param list
	 * @param filePath
	 * @throws Exception 
	 */
	public void saveExcel(List<LinkedHashMap<String, Object>> list, String filePath) throws Exception {

		Workbook workbook = null;
		FileOutputStream out = null; 

		try {
			workbook = createExcel(list);
			out = new FileOutputStream(filePath);		// 경로와 파일 생성
			workbook.write(out);						// 파일에 내용을 작성
		} catch (IOException ioe) {
			throw ioe;
		} finally {
			try {
				if (out != null) {
					out.close();
				}
			} catch (Exception e) {
				throw e;
			}
		}
	}

	/**
	 * 저장한 Excel을 Client에게 복사 다운로드
	 * @param list
	 * @param filePath
	 * @param newFileName
	 * @param response
	 * @throws Exception 
	 */
	public void copyExportExcel(List<LinkedHashMap<String, Object>> list, String filePath, String newFileName, HttpServletResponse response) throws Exception {

		File file = new File(filePath);											// 불러올 파일명

		InputStream is = null;
		OutputStream output = null;

		try {
			saveExcel(list, filePath);

			is = new FileInputStream(file);							

			String fileName = newFileName;										// 파일명
			fileName = new String(fileName.getBytes("UTF-8"), "ISO-8859-1");	// 파일명 한글 Encoding
			String fileExtension = ".xlsx";										// 확장자

			int bytelength = (int) file.length();
			byte fileByte[] = new byte[bytelength];

			response.reset(); // Some JSF component library or some Filter might have set some headers in the buffer beforehand. We want to get rid of them, else it may collide.
			response.setContentType("application/octet-stream"); // Check http://www.iana.org/assignments/media-types for all types. Use if necessary ServletContext#getMimeType() for auto-detection based on filename.
			response.setHeader("Content-Disposition", "attachment;filename=\"" + fileName + fileExtension + "\""); // The Save As popup magic is done here. You can give it any filename you want, this only won't work in MSIE, it will use current request URL as filename instead.

			// Write file to response.
			output = response.getOutputStream();

			for (int nChunk = is.read(fileByte); nChunk!=-1; nChunk = is.read(fileByte)) {
				output.write(fileByte);
			}
		} catch (IOException ioe) {
			throw ioe;
		} finally {
			try {
				if (is != null) {
					is.close();
				}
				if (output != null) {
					output.close();
				}
			} catch (Exception e) {
				throw e;
			}
		}
	}

	/**
	 * 생성한 Excel을 File로 생성하지 않고 Client에게 다운로드 
	 * @param list
	 * @param newFileName
	 * @param response
	 * @throws Exception 
	 */
	public void excelExport(List<LinkedHashMap<String, Object>> list, String newFileName, HttpServletResponse response) throws Exception {

		Workbook workbook = null; 
		OutputStream output = null;
		try {
			workbook = createExcel(list);

			String fileName = newFileName;										// 파일명
			fileName = new String(fileName.getBytes("UTF-8"), "ISO-8859-1");	// 파일명 한글 Encoding
			String fileExtension = ".xlsx";										// 확장자

			response.reset(); // Some JSF component library or some Filter might have set some headers in the buffer beforehand. We want to get rid of them, else it may collide.
			response.setContentType("application/octet-stream"); // Check http://www.iana.org/assignments/media-types for all types. Use if necessary ServletContext#getMimeType() for auto-detection based on filename.
			response.setHeader("Content-Disposition", "attachment;filename=\"" + fileName + fileExtension + "\""); // The Save As popup magic is done here. You can give it any filename you want, this only won't work in MSIE, it will use current request URL as filename instead.

			// Write file to response.
			output = response.getOutputStream();
			workbook.write(output);
		} catch (IOException ioe) {
			throw ioe;
		} finally {
			try {
				if (output != null) {
					output.close();
				}
			} catch (Exception e) {
				throw e;
			}
		}
	}

}
