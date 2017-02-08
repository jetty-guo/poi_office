package com.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.pdfbox.io.RandomAccessBufferedFileInputStream;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hslf.extractor.PowerPointExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;

public class TestDemo {

	public static void main(String[] args) throws Exception {

		try {
			String Xlscontent = getExcelAsFile("C:\\Users\\Administrator\\Downloads\\16050601.xls");
			//System.out.println(Xlscontent);

			String Wordcontent = getWorldAsFile("C:\\Users\\Administrator\\Downloads\\Inceptor-&-Hyperbase实例演示.docx");
		//	System.out.println(Wordcontent);
			
			String PPTcontent = getPPTAsFile("C:\\Users\\Administrator\\Downloads\\信道分配策略.ppt");
			System.out.println(PPTcontent);
			
			String PFDcontent = readPDF("C:\\Users\\Administrator\\Downloads\\ShopNC多用户商城系统平台手册.pdf");
			//	System.out.println(PFDcontent);

			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static String readPDF(String file) throws IOException {
		String result = null;
		FileInputStream is = null;
		PDDocument document = null;
		try {
			is = new FileInputStream(file);
			PDFParser parser = new PDFParser(new RandomAccessBufferedFileInputStream(is));
			parser.parse();
			document = parser.getPDDocument();
			PDFTextStripper stripper = new PDFTextStripper();
			result = stripper.getText(document);
		} finally {
			if (is != null) {
				is.close();
			}
			if (document != null) {
				document.close();
			}
		}
		return result;
	}

	private static String getPPTAsFile(String filepath) throws Exception {
		InputStream is = new FileInputStream(new File(filepath));
		PowerPointExtractor extractor = new PowerPointExtractor(is);
		extractor.close();
		return extractor.getText();
	}

	private static String getWorldAsFile(String filepath) {

		String content = null;
		try {
			String fileType = filepath.substring(filepath.lastIndexOf(".") + 1, filepath.length());

			if (fileType.equals("doc")) {
				InputStream is = new FileInputStream(new File(filepath));
				WordExtractor ex = new WordExtractor(is);
				String text2003 = ex.getText();
				System.out.println();
				content = text2003;
			} else if (fileType.equals("docx")) {
				OPCPackage opcPackage = POIXMLDocument.openPackage(filepath);
				POIXMLTextExtractor extractor = new XWPFWordExtractor(opcPackage);
				String text2007 = extractor.getText();
				content = text2007;

			} else {
				throw new Exception("读取的不是excel文件");
			}

		} catch (Exception e) {
			e.printStackTrace();
		}

		return content;
	}

	private static String getExcelAsFile(String filepath) throws Exception {
		String fileType = filepath.substring(filepath.lastIndexOf(".") + 1, filepath.length());
		InputStream is = null;
		Workbook wb = null;
		try {
			is = new FileInputStream(filepath);

			if (fileType.equals("xls")) {
				wb = new HSSFWorkbook(is);
			} else if (fileType.equals("xlsx")) {
				wb = new XSSFWorkbook(is);
			} else {
				throw new Exception("读取的不是excel文件");
			}

			StringBuilder sb = new StringBuilder();
			int sheetSize = wb.getNumberOfSheets();
			for (int i = 0; i < sheetSize; i++) {// 遍历sheet页
				Sheet sheet = wb.getSheetAt(i);

				int rowSize = sheet.getLastRowNum() + 1;
				for (int j = 0; j < rowSize; j++) {// 遍历行
					Row row = sheet.getRow(j);
					if (row == null) {// 略过空行
						continue;
					}
					int cellSize = row.getLastCellNum();// 行中有多少个单元格，也就是有多少列
					for (int k = 0; k < cellSize; k++) {
						Cell cell = row.getCell(k);
						String value = null;
						if (cell != null) {
							value = cell.toString();
						}
						sb.append(value);
					}
				}
			}

			return sb.toString();
		} catch (FileNotFoundException e) {
			throw e;
		} finally {
			if (wb != null) {
				wb.close();
			}
			if (is != null) {
				is.close();
			}
		}
	}
}
