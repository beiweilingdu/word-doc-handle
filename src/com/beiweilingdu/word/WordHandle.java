package com.beiweilingdu.word;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.UUID;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class WordHandle {
	
	private static final String USER_DIR = System.getProperty("user.dir");
	
	/**
	 * main
	 * @param args
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static void main(String[] args) throws FileNotFoundException, IOException {
		WordHandle wh = new WordHandle();
//		List<Integer> sections = new ArrayList<Integer>();
//		sections.add(1);
//		sections.add(3);
//		sections.add(11);
//		sections.add(12);
//		
//		wh.cutDocument("f:/13_12030084.docx", sections);
		
//		wh.scanIllegalWord("G:\\user\\wuxiaolin\\项目开发\\mspm\\产品手册文档示例\\产品资料--正式文件20141204(测试用)");
		wh.scanIllegalWord(args[0]);
	}
	
	/**
	 * 扫描某路径下所有不规范的word文档，包括word版本、内容段落不全等。
	 * 扫描的结果保存到txt文件中，内容是不规范文档的绝对路径和文件名
	 * @param dir
	 * @throws IOException 
	 */
	public synchronized void scanIllegalWord(String dir) throws IOException {
		
		/**
		 * 1、判断dir路径是否有效
		 * 2、递归扫描dir下所有文件，遇到.doc文件，就直接记录为不规范文档。并把.docx文件的绝对路径保存到一个List<String>中，等着下一步进行文档内容的扫描
		 * 3、循环处理每一个.docx文档，判断其内容是否规范
		 */
		// 参数校验
		if(dir == null || dir == "") {
			log("参数为空，请检查！");
			return;
		}
		// 路径校验
		File dirFile = new File(dir);
		if(!dirFile.exists()) {
			log("扫描路径不存在，请检查！");
			return;
		}
		
		// 递归扫描路径下所有文件
		List<String> result = new ArrayList<String>();
		this.scanAllFile(dir, result);
		
		// 判断内容是否规范
		List<String> keywords = this.getKeywordList();
		if(result != null) {
			
			File outfile = new File(USER_DIR + "/"+UUID.randomUUID()+".txt");
//			if(!outfile.exists()) {
//				outfile.mkdirs();
				outfile.createNewFile();
//			}
			BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outfile),"utf-8"));
			for(String filepath : result) {
				if(filepath.endsWith(".docx")) {
					File file = new File(filepath);
					if(file.exists()) {
						List<String> lostKeyword = new ArrayList<String>();
						Map<String, String> err = new HashMap<String, String>();
						XWPFDocument wordoc = new XWPFDocument(new FileInputStream(file));
						if(!this.isLegalWord(wordoc, keywords, lostKeyword, err)) {
							// TODO 记录文档内容不规范
							System.out.println("=========================================================");
							bw.write("=========================================================\r\n");
							bw.write(">文件路径：" + filepath + "\r\n");
							System.out.println(">文件路径：" + filepath);
							if(err.get("error") != null && err.get("error") != "") {
								bw.write(">格式错误：" + err.get("error") + "\r\n");
								System.out.println(">格式错误：" + err.get("error"));
							}
							
							if(lostKeyword.size() > 0) {
								bw.write(">缺失标签：\r\n");
								System.out.println(">缺失标签：");
								for(String lostkey : lostKeyword) {
									bw.write("\t" + lostkey + "\r\n");
									System.out.println("\t" + lostkey);
								}
							}
						} else {
							// TODO 规范的文档
//							System.out.println("================ " + filepath + " -> 通过 ===============");
						}
					}
				} else if(filepath.endsWith(".doc")) {
					// TODO 记录文档版本不规范
					bw.write("=========================================================\r\n");
					System.out.println("=========================================================");
					bw.write(">文件路径：" + filepath + "\r\n");
					System.out.println(">文件路径：" + filepath);
					bw.write(">版本错误：word文档版本不是2007及以上版本\r\n");
					System.out.println(">版本错误：word文档版本不是2007及以上版本");
				}
			}
			bw.write("=========================================================");
			bw.close();
			System.out.println("=========================================================");
			System.out.println(">检测结果已经保存到["+outfile.getAbsolutePath()+"]");
		}
	}
	
	/**
	 * 判断文档内容是否规范
	 * @param wordoc word文档
	 * @param keywords 关键字列表
	 * @param lostKeyword 丢失的关键字列表
	 * @return
	 */
	public Boolean isLegalWord(XWPFDocument wordoc, List<String> keywords, List<String> lostKeyword, Map<String, String> err) {
		
		List<XWPFTable> tables = wordoc.getTables();
		if(tables == null || tables.size() > 1) {
			err.put("error", "存在多个表格");
			return false;
		}
		
		// 读取文档表格
		XWPFTable table = tables.get(0);
		
		// 读取表格的行
		List<XWPFTableRow> rows = table.getRows();
		
		boolean isOK = true;
		
		for(String keyword : keywords) {
			boolean ok = false;
			// 遍历行
			for(XWPFTableRow row : rows) {
				// 某行第一个单元格
				XWPFTableCell cell = row.getCell(0);
				
				// 单元格内容
				String cellText = cell.getText();
				
				if(cellText != null && cellText.contains(keyword)) {
					// 当前关键字存在，继续下一个关键字的扫描
					ok = true;
					break;
				}
			}
			// 某一个关键字不存在，那么整个文档就是不OK的
			if(!ok) {
				lostKeyword.add(keyword);
				isOK = false;
			}
		}
		
		return isOK;
	}
	
	/**
	 * 递归扫描路径下所有文件
	 * @param dir
	 * @return
	 */
	public List<String> scanAllFile(String dir, List<String> result) {
		// 参数校验
		if(dir == null || dir == "") {
			log("参数为空，请检查！");
			return null;
		}
		// 路径校验
		File dirFile = new File(dir);
		if(!dirFile.exists()) {
			log("扫描路径不存在，请检查！");
			return null;
		}
		
		File[] fileList = dirFile.listFiles();
		
		for(File file : fileList) {
			if(file.isFile()) {
				result.add(file.getAbsolutePath());
			} else {
				scanAllFile(file.getAbsolutePath(), result);
			}
		}
		return result;
	}
	
	/**
	 * 获取关键词内容
	 * @return
	 * @throws IOException
	 */
	public List<String> getKeywordList() throws IOException {
		String keywords = this.getProperty("keywords");
		String[] keywordArr = keywords.split(",");
		return Arrays.asList(keywordArr);
	}
	
	/**
	 * 把原有word文档的指定部分保留下来，其他删掉
	 * @param filePath
	 * @param sections
	 * @throws IOException 
	 */
	public void cutDocument(String filePath, List<Integer> sections) throws IOException {
		// 目标文件路径
		String desFilePath = this.generateNewFilename(filePath);
		
		// 把word文档读取到内存
		XWPFDocument wordoc = new XWPFDocument(new FileInputStream(new File(filePath)));
		
		// 获取关键词内容
		List<String> keywordList = this.getKeywordList();
		
		// 读取文档表格
		XWPFTable table = wordoc.getTables().get(0);
		
		// 读取表格的行
		List<XWPFTableRow> rows = table.getRows();
		
		// 记录要删除的行编号
		List<Integer> rowtodel = new ArrayList<Integer>();
		
		// 遍历行
		for(XWPFTableRow row : rows) {
			// 某行第一个单元格
			XWPFTableCell cell = row.getCell(0);
			
			// 单元格内容
			String cellText = cell.getText();
			
			// 遍历关键词内容
			for(String keyword : keywordList) {
				// 当前单元格内容是否包含了关键词内容，如果是，则该单元格所在行数就是所定义的section，需要继续往下处理这个section是否要删除
				if(cellText.contains(keyword)) {
					// 关键词内容在文档中的section位置
					int sectionNum = keywordList.indexOf(keyword) + 1;
					// 现在所处的行编号，和数组索引类似，从0计算
					int rowNum = rows.indexOf(row);
					
					// 记录一下当前读取的内容
					System.out.println("找到第[" + sectionNum + "]部分\t行号->[" +rowNum + "]\t关键字内容->[" + keyword + "]");
					
					/* 
					 * sections是用户想要保留的section集合，当前所在section的编号是sectionNum，
					 * 如果当前section不在用户要求范围内，则记录到标记列表中（前面定义的rowtodel）
					 */
					 
					if(!sections.contains(sectionNum)) {
						// 标记当前行
						rowtodel.add(rowNum);

						// 移到下一行
						rowNum++;
						
						// 从下一行开始逐行标记
						for(int curRowNum = rowNum; curRowNum < rows.size(); curRowNum++ ) {
							// 如果这是最后一个section，那么就把剩余行都标记
							if(sectionNum >= keywordList.size()) {
								rowtodel.add(curRowNum);
//								break;
							} else {
								// 如果行的第一个单元格内容是某一个关键词，那么停止标记
								String tempText = table.getRow(curRowNum).getCell(0).getText();
								if(tempText.contains(keywordList.get(sectionNum))) {
									break;
								}
								// 标记
								rowtodel.add(curRowNum);
							}
						}
					}
				}
			}
		}
		
		// 记录一下要删除的行
		System.out.println("\n要删除的行" + rowtodel.toString());
		
		// 删除标记的行
		for(int i = 0; i < rowtodel.size(); i++) {
			int curRowNum = rowtodel.get(i);
			table.removeRow(curRowNum);
			
			// 删除一行后，该行以后的行的编号都要减少一行
			for(int j = i + 1; j < rowtodel.size(); j++) {
				rowtodel.set(j, rowtodel.get(j) - 1);
			}
		}
		
		// 上面的操作只是影响内存中的数据，一定要写入文件
		wordoc.write(new FileOutputStream(new File(desFilePath)));
	}
	
	/**
	 * 生成新文件路径
	 * @param originalFilename
	 * @return
	 */
	public String generateNewFilename(String originalFilename) {
		Date date = new Date();
		int dotIndex = originalFilename.lastIndexOf(".");
		String extName = originalFilename.substring(dotIndex + 1);
		String pathWithoutExt = originalFilename.substring(0, dotIndex);
		String formatDateStr = String.format("%tY%tm%td%tH%tM%tS%tz%tZ", date, date, date, date, date, date, date, date);
		return pathWithoutExt + "_" + formatDateStr + "." + extName;
	}
	
	/**
	 * 获取配置属性
	 * @param key
	 * @return
	 * @throws IOException
	 */
	public String getProperty(String key) throws IOException {
		Properties prop = new Properties();
		InputStream inputStream = WordHandle.class.getClassLoader().getResourceAsStream("config.properties");
		prop.load(inputStream);
		inputStream.close();
		return prop.getProperty(key);
	}	
	
	/**
	 * 简单日志打印
	 * @param message
	 */
	public void log(String message) {
		System.out.println("["+(new Date())+"]" + message);
	}

}
