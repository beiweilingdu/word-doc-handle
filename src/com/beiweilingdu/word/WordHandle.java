package com.beiweilingdu.word;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import org.apache.poi.hdf.extractor.WordDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class WordHandle {
	
	/**
	 * main
	 * @param args
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static void main(String[] args) throws FileNotFoundException, IOException {
		WordHandle wh = new WordHandle();
		List<Integer> sections = new ArrayList<Integer>();
		sections.add(1);
		sections.add(3);
		sections.add(11);
		sections.add(12);
		
		wh.cutDocument("f:/13_12030084.docx", sections);
		
//		wh.test();
	}
	
	public void test() throws FileNotFoundException, IOException {
		XWPFDocument wordoc = new XWPFDocument(new FileInputStream(new File("f:/13_12030084.docx")));
		
		XWPFTable table = wordoc.getTables().get(0);
		
		XWPFTableRow row = table.getRow(0);
		
		List<XWPFTableCell> cells = row.getTableCells();
		
		System.out.println(cells.size());
		
		row.removeCell(5);
		
		table.addRow(row, 0);
		
		table.removeRow(1);
		
		System.out.println(row.getTableCells().size());
		
//		System.out.println(row.getCell(0).getText());
//		System.out.println(row.getCell(1).getText());
//		System.out.println(row.getCell(2).getText());
//		System.out.println(row.getCell(3).getText());
//		System.out.println(row.getCell(4).getText());
//		System.out.println(row.getCell(5).getText());
		
		
		wordoc.write(new FileOutputStream(new File("f:/13_12030084_test.docx")));
		
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
}
