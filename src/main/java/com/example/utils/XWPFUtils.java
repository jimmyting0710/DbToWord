package com.example.utils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

public class XWPFUtils {

	/**
	 * 打开word文档
	 * 
	 * @param path 文档所在路径
	 * @return
	 * @throws IOException
	 * @Author Huangxiaocong 2018年12月1日 下午12:30:07
	 */
	public XWPFDocument openDoc(InputStream is) throws IOException {
		return new XWPFDocument(is);
	}

	/**
	 * 替換段落裡面的變數
	 *
	 * @param doc    要替換的文件
	 * @param params 引數
	 */
	public static void replaceInPara(XWPFDocument doc, Map<String, Object> params) {
		Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
		XWPFParagraph para;
		while (iterator.hasNext()) {
			para = iterator.next();
			replaceInPara(para, params);
		}
	}

	/**
	 * 替換段落裡面的變數
	 *
	 * @param para   要替換的段落
	 * @param params 引數
	 */
	private static void replaceInPara(XWPFParagraph para, Map<String, Object> params) {
		Matcher matcher = matcher(para.getParagraphText());
		String runText = "";
		String keyString;
		if (matcher.find()) {
			runText = getSpliString(para);
			while ((matcher = matcher(runText)).find()) {
				keyString = matcher.group(0);
				runText = matcher.replaceFirst(Matcher.quoteReplacement(String.valueOf(params.get(keyString))));
			}
			// 直接呼叫XWPFRun的setText()方法設定文字時，在底層會重新建立一個XWPFRun，把文字附加在當前文字後面，
			// 所以我們不能直接設值，需要先刪除當前run,然後再自己手動插入一個新的run。
			para.insertNewRun(0).setText(runText);
		}

	}

	public void replaceTable(XWPFDocument doc, String tagString, List<Map<String, String>> dataList,
			List<String> queryColArray) {
//		for(Map<String, String> dd:dataList) {
//			for(Entry<String, String> aa:dd.entrySet()) {
//				System.out.print(aa.getKey() + "    ");
//				System.out.println(aa.getValue());
//			}
//		}
		
		List<XWPFParagraph> paras = doc.getParagraphs();
		for (XWPFParagraph para : paras) {
			//
			String runString = para.getParagraphText();
			// Match Paragraph Test
			Matcher matcher = matcher(para.getParagraphText());
			while (matcher.find()) {
				runString = matcher.group(0);
			}
			List<XWPFRun> runs = para.getRuns();
//				String runString = run.getText(0).trim();
			if (runString != null) {
				if (runString.indexOf(tagString) >= 0) {
					for (XWPFRun run : runs) {
						run.setText(runString.replace(tagString, ""), 0);
					}
					XmlCursor cursor = para.getCTP().newCursor();
					XWPFTable table = doc.insertNewTbl(cursor);

					// 自設表格寬度，不設定會auto
//					CTTblPr tablePr = table.getCTTbl().getTblPr();
//					CTTblWidth width = tablePr.addNewTblW();
//					width.setW(BigInteger.valueOf(8500));
					fillHeaderData(table, dataList, queryColArray);
					fillTableData(table, dataList, queryColArray);
					setTableLocation(table, "center");
				}
			}
		}
	}

	public static String getSpliString(XWPFParagraph para) {
		List<XWPFRun> runs;
		runs = para.getRuns();
		StringBuilder runText = new StringBuilder();
		if (runs.size() > 0) {
			int j = runs.size();
			for (int i = 0; i < j; i++) {
				XWPFRun run = runs.get(0);
				String i1 = run.toString();
				runText.append(i1);
				para.removeRun(0);

			}

		}
		return runText.toString();
	}

	/**
	 * 正則匹配字串
	 *
	 * @param str
	 * @return
	 */
	public static Matcher matcher(String str) {
		Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}

	
	


	public void fillHeaderData(XWPFTable table, List<Map<String, String>> tableData, List<String> HeaderName) {
		XWPFTableRow headerRow = table.getRow(0);
		XWPFTableCell cell;
		
		for (int i = 0; i < HeaderName.size(); i++) {
//			if(HeaderName.get(i).equals("DSN")) {   //不印dsn
//				continue;
//			}
			if (headerRow.getCell(i) == null) {
				cell = headerRow.createCell();
			} else {
				cell = headerRow.getCell(i);
			}
			XWPFParagraph cellParagraph = cell.getParagraphArray(0);
			XWPFRun cellParagraphRun = cellParagraph.createRun();
			cellParagraphRun.setText(HeaderName.get(i));

		}
	}

	/**
	 * 往表格中填充数据
	 * 
	 * @param table
	 * @param tableData
	 * @Author Huangxiaocong 2018年12月16日
	 */
	public void fillTableData(XWPFTable table, List<Map<String, String>> tableData, List<String> HeaderName) {
		for (int i = 0; i < tableData.size(); i++) {
			XWPFTableRow row = table.createRow();
			Map<String, String> item = tableData.get(i);
			for (int j = 0; j < row.getTableCells().size(); j++) {
				XWPFTableCell cell = row.getCell(j);
				XWPFParagraph cellParagraph = cell.getParagraphArray(0);
				XWPFRun cellParagraphRun = cellParagraph.createRun();
				cellParagraphRun.setText(item.get(HeaderName.get(j)));
			}
		}
	}

	/**
	 * 設定表格位置
	 *
	 * @param xwpfTable
	 * @param location  整個表格居中center,left居左，right居右，both兩端對齊
	 */
	public static void setTableLocation(XWPFTable xwpfTable, String location) {
		CTTbl cttbl = xwpfTable.getCTTbl();
		CTTblPr tblpr = cttbl.getTblPr() == null ? cttbl.addNewTblPr() : cttbl.getTblPr();
		CTJc cTJc = tblpr.addNewJc();
		cTJc.setVal(STJc.Enum.forString(location));
	}




	

	// 自已後來找的
	public static void searchAndReplace(XWPFDocument document, Map<String, String> data1) {
		try {

			/**
			 * 替換表格中的指定文字
			 */
			Iterator<XWPFTable> itTable = document.getTablesIterator(); // 獲得Word的表格

			while (itTable.hasNext()) { // 遍歷表格
				XWPFTable table = (XWPFTable) itTable.next();
				int count = table.getNumberOfRows(); // 獲得表格總行數
				for (int i = 0; i < count; i++) { // 遍歷表格的每一行
					XWPFTableRow row = table.getRow(i); // 獲得表格的行
					List<XWPFTableCell> cells = row.getTableCells(); // 在行元素中，獲得表格的單元格
					for (XWPFTableCell cell : cells) { // 遍歷單元格
						for (Entry<String, String> e : data1.entrySet()) {
							if (cell.getText().equals(e.getKey())) { // 如果單元格中的變數和‘鍵’相等，就用‘鍵’所對應的‘值’代替。
								cell.removeParagraph(0); // 所以這裡就要求每一個單元格只能有唯一的變數。
								cell.setText((String) e.getValue());
							}
						}
					}
				}
			}
//			FileOutputStream outStream = null;
//            outStream = new FileOutputStream(destPath);
//            document.write(outStream);
//            outStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}