package com.example.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;
import java.util.Properties;
import java.util.Set;
import java.util.TreeSet;
import java.util.stream.Collectors;

import javax.swing.plaf.synth.SynthOptionPaneUI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.example.utils.XWPFUtils;

import fr.opensagres.xdocreport.core.utils.StringUtils;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;

public class ToWord {

	private final static Logger logger = LoggerFactory.getLogger(ToWord.class);
	List<String> queryColArray;// 要抓取的欄位
	List<String> tableOutputKey;// 範本需要抓取的欄位
	List<String> table3OutputKey;// 範本output需要抓取的欄位
	List<String> table4OutputKey;// 範本output需要抓取的欄位
	List<String> table5OutputKey;// 範本output需要抓取的欄位
	String tempFileFolderPath; // word範本位置
	String destFileFolderPath; // word輸出資料夾位置
	XWPFUtils XWPFUtils = new XWPFUtils();
	Map<Integer, String> HeaderName = new HashMap<Integer, String>();

	public ToWord() {
		Properties pro = new Properties();
		// 設定檔位置
		String config = "config.properties";
		try {
			System.out.println("DbToWord 執行");
			// 讀取設定檔
			logger.info(MessageFormat.format("設定檔位置: {0}", config));
			pro.load(new FileInputStream(config));
			// 讀取資料夾位置
			tempFileFolderPath = pro.getProperty("tempFile");
			destFileFolderPath = pro.getProperty("destFile");
			// 讀取需要抓取的欄位名稱
			tableOutputKey = Arrays.asList(pro.getProperty("tableOutputKey").split(","));
			table3OutputKey = Arrays.asList(pro.getProperty("table3OutputKey").split(","));
			table4OutputKey = Arrays.asList(pro.getProperty("table4OutputKey").split(","));
			table5OutputKey = Arrays.asList(pro.getProperty("table5OutputKey").split(","));
		} catch (FileNotFoundException e) {
			logger.info(e.toString());
			e.printStackTrace();
		} catch (IOException e) {
			logger.info(e.toString());
			e.printStackTrace();
		} catch (Exception e) {
			logger.info(e.toString());
			e.printStackTrace();
		}
	}

	/**
	 * 將資料重組並輸出Word
	 * 
	 * @param input
	 * @param output
	 * @param data
	 * @param catalog
	 * 
	 * @param excelDataList 整理過的Excel檔案
	 */
	
	
	// 聯合資料 db直接篩選完
	public void outPutToWork(List<Map<String, String>> input, List<Map<String, String>> output,
			List<Map<String, String>> data, List<Map<String, String>> catalog) {
		catalog.addAll(output);

		// 把key塞進headername
		int i = 1;
		for (Map<String, String> all : catalog) {
			for (Entry<String, String> findheader : all.entrySet()) {
				HeaderName.put(i++, findheader.getKey());
			}
			break;
		}

	
		// 抓出不重複的JCL
		HashSet<String> jclKeys = new HashSet<>();
		catalog.forEach(cn -> {
			jclKeys.add(cn.get("JCL"));
		});
//		HashSet<String> adKeys = new HashSet<>();
//		catalog.forEach(cn -> {
//			adKeys.add(cn.get("AD"));
//		});

//		for (String b : adKeys) {
//			System.out.println(b);
//		}

		int fileCount = 0;
//		logger.info("預計產出 {} 個檔案", jclKeys.size());
		// 將不重複的相同JCL_NAME的資料Group to List並輸出word

		for (String classKey : jclKeys) {
			logger.info("開始輸出:" + classKey);
//				if (classKey.equals("JBP02C")) {
//					Map<String, List<Map<String, String>>> toWordList;
//					toWordList = data.stream().filter(dbModel -> classKey.equals(dbModel.get("JCL")))
//							.collect(Collectors.groupingBy(d -> d.get("AD"))); // 篩選classkey之後回傳
//					for (Entry<String, List<Map<String, String>>> a : toWordList.entrySet()) {

			List<Map<String, String>> toWordList;
			toWordList = catalog.stream().filter(dbModel -> classKey.equals(dbModel.get("JCL")))
					.collect(Collectors.toList()); // 篩選classkey之後回傳

			// 把ad分組
			Map<String, List<Map<String, String>>> ad;
			ad = toWordList.stream().collect(Collectors.groupingBy(d -> d.get("AD")));
			for (Entry<String, List<Map<String, String>>> a : ad.entrySet()) {
				// 輸出Word
//				createWord(a.getValue());  //暫時封住
//						createWord(toWordList);
				logger.info("已輸出 JCL Name: " + classKey);
//					// 輸出完之後，刪除，節省資源。
				toWordList.forEach(Item -> catalog.remove(Item));
				toWordList.clear();
				fileCount++;
			}
		}

//		}
		logger.info("實際產出 {} 個檔案", fileCount);
	}

	// 單一資料
	public void outPutToWork1(List<Map<String, String>> data) {

		// 把key塞進headername
		int i = 1;
		for (Map<String, String> all : data) {
			for (Entry<String, String> findheader : all.entrySet()) {
				HeaderName.put(i++, findheader.getKey());
			}
			break;
		}

		// 抓出不重複的JCL
		HashSet<String> jclKeys = new HashSet<>();
		data.forEach(cn -> {
			jclKeys.add(cn.get("JCL"));
		});
//		HashSet<String> adKeys = new HashSet<>();
//		catalog.forEach(cn -> {
//			adKeys.add(cn.get("AD"));
//		});

//		for (String b : adKeys) {
//			System.out.println(b);
//		}

		int fileCount = 0;
//		logger.info("預計產出 {} 個檔案", jclKeys.size());
		// 將不重複的相同JCL_NAME的資料Group to List並輸出word

		for (String classKey : jclKeys) {
			logger.info("開始輸出:" + classKey);
//				if (classKey.equals("JBP02C")) {
//					Map<String, List<Map<String, String>>> toWordList;
//					toWordList = data.stream().filter(dbModel -> classKey.equals(dbModel.get("JCL")))
//							.collect(Collectors.groupingBy(d -> d.get("AD"))); // 篩選classkey之後回傳
//					for (Entry<String, List<Map<String, String>>> a : toWordList.entrySet()) {

			List<Map<String, String>> toWordList;
			toWordList = data.stream().filter(dbModel -> classKey.equals(dbModel.get("JCL")))
					.collect(Collectors.toList()); // 篩選classkey之後回傳

			// 把ad分組
			Map<String, List<Map<String, String>>> ad;
			ad = toWordList.stream().collect(Collectors.groupingBy(d -> d.get("AD")));
			for (Entry<String, List<Map<String, String>>> a : ad.entrySet()) {
				// 輸出Word
//				createWord(a.getValue());   //暫時封住
//						createWord(toWordList);
				logger.info("已輸出 JCL Name: " + classKey);
//					// 輸出完之後，刪除，節省資源。
				toWordList.forEach(Item -> data.remove(Item));
				toWordList.clear();
				fileCount++;
			}
		}

//		}
		logger.info("實際產出 {} 個檔案", fileCount);
	}

	// 雙資料，把data2丟到createword合併
	public void outPutToWork2(List<Map<String, String>> data, List<Map<String, String>> data2) {

		// 把key塞進headername
		int i = 1;
		for (Map<String, String> all : data) {
			for (Entry<String, String> findheader : all.entrySet()) {
				HeaderName.put(i++, findheader.getKey());
			}
			break;
		}

		// 抓出不重複的JCL
		HashSet<String> jclKeys = new HashSet<>();
		data.forEach(cn -> {
			jclKeys.add(cn.get("JCL"));
		});

		int fileCount = 0;
		// 將不重複的相同JCL_NAME的資料Group to List並輸出word

		for (String classKey : jclKeys) {
			logger.info("開始輸出:" + classKey);

			List<Map<String, String>> toWordList;
			toWordList = data.stream().filter(dbModel -> classKey.equals(dbModel.get("JCL")))
					.collect(Collectors.toList()); // 篩選classkey之後回傳

			// 把ad分組
			Map<String, List<Map<String, String>>> ad;
			ad = toWordList.stream().collect(Collectors.groupingBy(d -> d.get("AD")));
			for (Entry<String, List<Map<String, String>>> a : ad.entrySet()) {
				// 輸出Word
				createWord(a.getValue(), data2);
//						createWord(toWordList);
				logger.info("已輸出 JCL Name: " + classKey);
//					// 輸出完之後，刪除，節省資源。
				toWordList.forEach(Item -> data.remove(Item));
				toWordList.clear();
				fileCount++;
			}
		}
		logger.info("實際產出 {} 個檔案", fileCount);
	}

	// 雙資料
	public void outPutToWork3(List<Map<String, String>> data,List<Map<String, String>> db2data,List<Map<String, String>> imsdbdata) {
		
		
		// 把key塞進headername
		int i = 1;
		for (Map<String, String> all : data) {
			for (Entry<String, String> findheader : all.entrySet()) {
				HeaderName.put(i++, findheader.getKey());
			}
			break;
		}
		
		// 抓出不重複的JCL
		HashSet<String> jclKeys = new HashSet<>();
		data.forEach(cn -> {
			jclKeys.add(cn.get("JCL"));
		});

		int fileCount = 0;
		// 將不重複的相同JCL_NAME的資料Group to List並輸出word

		for (String classKey : jclKeys) {
			logger.info("開始輸出:" + classKey);

			List<Map<String, String>> toWordList;
			toWordList = data.stream().filter(dbModel -> classKey.equals(dbModel.get("JCL")))
					.collect(Collectors.toList()); // 篩選classkey之後回傳

			// 把ad分組
			Map<String, List<Map<String, String>>> ad;
			ad = toWordList.stream().collect(Collectors.groupingBy(d -> d.get("AD")));
			for (Entry<String, List<Map<String, String>>> a : ad.entrySet()) {
				// 輸出Word
				createWord1(a.getValue(),db2data,imsdbdata);
//						createWord(toWordList);
				logger.info("已輸出 JCL Name: " + classKey);
//					// 輸出完之後，刪除，節省資源。
				toWordList.forEach(Item -> data.remove(Item));
				toWordList.clear();
				fileCount++;
			}
		}
		logger.info("實際產出 {} 個檔案", fileCount);
	}
	
	
	
	
	
//	D:\si1204\Desktop\單元測試個案\Batch
	public void createWord(List<Map<String, String>> data, List<Map<String, String>> data2) {

		if (!new File(destFileFolderPath + "/" + data.get(0).get("SYSTEM_OPERATION") + "/" + data.get(0).get("AD"))
				.exists()) {
			new File(destFileFolderPath + "/" + data.get(0).get("SYSTEM_OPERATION") + "/" + data.get(0).get("AD"))
					.mkdirs();
		}
		try (InputStream is = new FileInputStream(tempFileFolderPath);
				OutputStream os = new FileOutputStream(
						destFileFolderPath + "/" + data.get(0).get("SYSTEM_OPERATION") + "/" + data.get(0).get("AD")
								+ "/" + data.get(0).get("AD") + "_" + data.get(0).get("JCL") + ".docx");) {

			XWPFDocument doc = XWPFUtils.openDoc(is);
			List<XWPFParagraph> xwpfParas = doc.getParagraphs();
//篩選完自建表格
//			List<Map<String, String>> Catalog = data.stream()
//					.filter(item-> (!"Y                                                 "
//									.equals(String.valueOf(item.get("PASSFORM")))
//									|| !"Y                                                 "
//											.equals(String.valueOf(item.get("IOCHECKLIST")))
//									||! "Y                                                 "
//											.equals(String.valueOf(item.get("ALLJCL")))
//									|| !"Y                                                 "
//											.equals(String.valueOf(item.get("SENDFILE")))))
//					.collect(Collectors.toList());
			List<Map<String, String>> Catalog = data.stream()
					.collect(Collectors.toList());
     
//			List<Map<String, String>> inputList = data.stream()
//					.filter(item -> "I".equals(String.valueOf(item.get("OPEN_MODE")))
//							|| "IO".equals(String.valueOf(item.get("OPEN_MODE")))
//							|| "SORTIN".equals(String.valueOf(item.get("DD")))
//					)
//					.collect(Collectors.toList());
//					.collect(Collectors.collectingAndThen(
//							Collectors.toCollection(() -> new TreeSet<>(Comparator.comparing(m -> m.get("DSN")))),
//							ArrayList::new));

//			List<Map<String, String>> ouputList = data.stream()
//					.filter(item -> ("O".equals(String.valueOf(item.get("OPEN_MODE")))
//							|| "IO".equals(String.valueOf(item.get("OPEN_MODE"))))
////							|| "SORTOUT".equals(String.valueOf(item.get("DD")))
//							&& ("Y                                                 "
//									.equals(String.valueOf(item.get("PASSFORM")))
//									|| "Y                                                 "
//											.equals(String.valueOf(item.get("IOCHECKLIST")))
//									|| "Y                                                 "
//											.equals(String.valueOf(item.get("ALLJCL")))
//									|| "Y                                                 "
//											.equals(String.valueOf(item.get("SENDFILE")))))
//					.collect(Collectors.toList());
			List<Map<String, String>> ouputList = data.stream()
					.filter(item -> "I-O".equals(String.valueOf(item.get("OPEN_MODE")))
							|| "INPUT,OUTPUT".equals(String.valueOf(item.get("OPEN_MODE")))
							|| "I-O,OUTPUT".equals(String.valueOf(item.get("OPEN_MODE")))
							|| "OUTPUT".equals(String.valueOf(item.get("OPEN_MODE")))
							|| "I-O,INPUT".equals(String.valueOf(item.get("OPEN_MODE"))))
					.collect(Collectors.toList());

			for (XWPFParagraph xwpfParagraph : xwpfParas) {
				String itemText = xwpfParagraph.getText();
				switch (itemText) {
				case "${catalog}":
					XWPFUtils.replaceTable(doc, itemText, Catalog, tableOutputKey);
					break;

//				case "${dataTable}":
//					XWPFUtils.replaceTable(doc, itemText, inputList, tableOutputKey);
//					break;

				case "${resultTable}":
					XWPFUtils.replaceTable(doc, itemText, ouputList, table3OutputKey);
					break;
				}
			}

			Map<String, Object> titledata = new HashMap<>();

			titledata.put("${SYSTEM_OPERATION}", data.get(0).get("SYSTEM_OPERATION"));
			titledata.put("${AD}", data.get(0).get("AD"));
			titledata.put("${JCL}", data.get(0).get("JCL"));
		
			// 取代資料
			XWPFUtils.replaceInPara(doc, titledata);

			Map<String, String> data1 = new HashMap<>();

			data1.put("${para}", data.get(0).get("PARA"));
			// 為情境更改 新增Data2
			Set<String> db2set = new HashSet<String>();
			Set<String> imsset = new HashSet<String>();
			String db2_include = null;
			String ims_get = null;
			for (Map<String, String> aa : data2) {
				if (((aa.get("AD").toString()).equals(data.get(0).get("AD").toString())
						|| (aa.get("AD").toString() == data.get(0).get("AD").toString()))
						&& ((aa.get("JCL").toString()).equals(data.get(0).get("JCL").toString())
								|| (aa.get("JCL").toString() == data.get(0).get("JCL").toString()))) {
					if (aa.get("DB2_INCLUDE") != null) {
						for (String s : aa.get("DB2_INCLUDE").split(",")) {
							db2set.add(s);
						}
						db2_include = String.join(",", db2set);
						data1.put("${db2_include}", db2_include);
					}
					if (aa.get("IMS_GET") != null) {
						for (String s : aa.get("IMS_GET").split(",")) {
							imsset.add(s);
						}
						ims_get = String.join(",", imsset);
						data1.put("${ims_get}", ims_get);
					}
				
					break;
				}else {
					data1.put("${db2_include}", " ");
					data1.put("${ims_get}", " ");
				}
			}

			XWPFUtils.searchAndReplace(doc, data1);
			doc.write(os);
		} catch (FileNotFoundException e) {
			logger.info(e.toString());
			e.printStackTrace();
		} catch (IOException e) {
			logger.info(e.toString());
			e.printStackTrace();
		}
	}

	
	public void createWord1(List<Map<String, String>> data , List<Map<String, String>> data2, List<Map<String, String>> data3) {

		if (!new File(destFileFolderPath + "/" + data.get(0).get("SYSTEM_OPERATION") + "/" + data.get(0).get("AD"))
				.exists()) {
			new File(destFileFolderPath + "/" + data.get(0).get("SYSTEM_OPERATION") + "/" + data.get(0).get("AD"))
					.mkdirs();
		}
		try (InputStream is = new FileInputStream(tempFileFolderPath);
				OutputStream os = new FileOutputStream(
						destFileFolderPath + "/" + data.get(0).get("SYSTEM_OPERATION") + "/" + data.get(0).get("AD")
								+ "/" + data.get(0).get("AD") + "_" + data.get(0).get("JCL") + ".docx");) {

			XWPFDocument doc = XWPFUtils.openDoc(is);
			List<XWPFParagraph> xwpfParas = doc.getParagraphs();
//篩選完自建表格

			List<Map<String, String>> Catalog = data.stream()
					.collect(Collectors.toList());
			
			List<Map<String, String>> db2 =data2.stream()
					.filter(item-> (data.get(0).get("AD")).equals(item.get("AD"))
							&& data.get(0).get("JCL").equals(item.get("JCL")))
					.collect(Collectors.toList());
			
			List<Map<String, String>> imsdb =data3.stream()
					.filter(item-> (data.get(0).get("AD")).equals(item.get("AD"))
							&& data.get(0).get("JCL").equals(item.get("JCL")))
					.collect(Collectors.toList());

			List<Map<String, String>> ouputList = data.stream() 
					.filter(item -> "IO".equals(String.valueOf(item.get("OPEN_MODE")))
							|| "OUTPUT".equals(String.valueOf(item.get("OPEN_MODE"))))
					.collect(Collectors.toList());

			for (XWPFParagraph xwpfParagraph : xwpfParas) {
				String itemText = xwpfParagraph.getText();
				switch (itemText) {
				case "${catalog}":
					XWPFUtils.replaceTable(doc, itemText, Catalog, tableOutputKey);
					break;

				case "${db2}":
					XWPFUtils.replaceTable(doc, itemText, db2, table4OutputKey);
					break;

				case "${imsdb}":
					XWPFUtils.replaceTable(doc, itemText, imsdb, table5OutputKey);
					break;	
					
				case "${resultTable}":
					XWPFUtils.replaceTable(doc, itemText, ouputList, table3OutputKey);
					break;
				}
			}

			Map<String, Object> titledata = new HashMap<>();

			titledata.put("${SYSTEM_OPERATION}", data.get(0).get("SYSTEM_OPERATION"));
			titledata.put("${AD}", data.get(0).get("AD"));
			titledata.put("${JCL}", data.get(0).get("JCL"));
		
			// 取代資料
			XWPFUtils.replaceInPara(doc, titledata);

			Map<String, String> data1 = new HashMap<>();

			data1.put("${para}", data.get(0).get("PARA"));
			

			XWPFUtils.searchAndReplace(doc, data1);
			doc.write(os);
		} catch (FileNotFoundException e) {
			logger.info(e.toString());
			e.printStackTrace();
		} catch (IOException e) {
			logger.info(e.toString());
			e.printStackTrace();
		}
	}
}
