package com.example;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import com.example.XWPFHandler.WordTemplate;

public class MainWordTemplate {
	
	private static WordTemplate template;
	private static final String DOCX_MODEL_PATH = "doc/document.docx";
	private static final String DOCX_FILE_WRITE = "D:/generate-document.docx";
	//private static final String DOCX_FILE_WRITE = "doc/generate-document.docx";
	
	public static void ExportDocument(Map<String, String> map){
		
		File file = new File(DOCX_MODEL_PATH);
		FileInputStream fileInputStream = null;
		try {
			fileInputStream = new FileInputStream(file);
			template = new WordTemplate(fileInputStream);
		} catch(IOException exception){
			exception.printStackTrace();
		}
		
		template.replaceTag(map);
		
		File file1 = new File(DOCX_FILE_WRITE);
		FileOutputStream out;
		try {
			out = new FileOutputStream(file1);
			BufferedOutputStream bos = new BufferedOutputStream(out);
			template.write(bos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		System.out.println("成功！");
	}
	
	public static void main(String[] args) {
		
		Map<String, String> map = new HashMap<String, String>();
		//设置参数
		map.put("title1", "标题一");
		map.put("title2", "标题二");
		map.put("cell1", "content");
		map.put("cell2", "content");
		map.put("cell3", "content");
		map.put("cell4", "content");
		map.put("year", "2018");
		map.put("month", "12");
		map.put("day", "27");
		map.put("test", "测试内容");
		
		ExportDocument(map);
		
	}
}
