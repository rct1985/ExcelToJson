package xls2json;

import xls2json.Base64Util;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

class FileAndMd5Class {
	public FileAndMd5Class(){
		fileName = "";
		md5CheckCode = "";
	}
	public String fileName;
	public String md5CheckCode;
}

public class Entrance {
	// >>>>Excel转Json相关
	public static final String MyPath = "";
	public static final String confFile = "conf.txt";
	public static boolean s_bIsWithMainKey = false;
	public static ArrayList<XLSStructure> xlsFileList;
	public static String s_jsonOutputPath;
	public static String xlsInputPath;
	public static String versionString = "=====xls2Json Version 20130620=====";

	// >>>>加密校验相关
	public static final String sf_strTempDir = "./tempDir/";
	public static final String sf_strCurrentDir = "./";
	public static final String sf_strInserterLabel = ">>>md5CheckCode begin >=-Rct-=<<<<<<"; // 插入点的位置
	public static boolean s_bIsEncrypt = false; // 是否加密
	public static boolean s_bIsMd5Check = true; // 是否md5校验, 这里不能控制,肯定会产生,
												// 但可以在项目代码里选择不用它
	public static String s_globalDataFileName; // 存放md5校验码及是否加密宏的文件
	public static ArrayList<FileAndMd5Class> s_strMd5Array;
	
	
	// inner class

	public static void main(String[] args) {
		xlsFileList = new ArrayList<XLSStructure>();
		System.out.println(versionString);
		processConvert();
		processAfterJson();
	}

	public static void processConvert() {
		File l_excuteFile = new File("./Excel2Json.jar");		

		SimpleDateFormat l_formatter = new SimpleDateFormat("yyyy年MM月dd日 HH:mm:ss");   
		Date l_date = new Date(l_excuteFile.lastModified());//获取最后更改时间
		String l_strDate = l_formatter.format(l_date);   		
		see("工具更新日期: " + l_strDate);
		
		see("Processing...");
		xlsFileList.clear();
		File file = new File(MyPath + confFile);
		BufferedReader reader = null;
		try {
			System.out.println("Reading lines...");
			reader = new BufferedReader(new FileReader(file));
			String tempString = null;
			int line = 1;
			XLSStructure tempStructure = null;
			while ((tempString = reader.readLine()) != null) {
				//System.out.println("line:" + line + ":" + tempString);
				System.out.println(tempString);
				tempString = tempString.trim();
				if (!tempString.isEmpty()) {
					// >>>>>>>>>>>>>>>>>>>配置选项<<<<<<<<<<<<<<<<<<
					//注释
					if(tempString.startsWith("/*") || tempString.startsWith("*")){
						//nothing...
					}
					// Excel的文件路径
					else if (tempString.startsWith("Input_Path_Excel")) {
						int l_iStartIndex = tempString.indexOf(":") + 1;
						xlsInputPath = tempString.substring(l_iStartIndex).trim();
					}
					// 导出的文件路径
					else if (tempString.startsWith("Output_Path_Json")) {
						int l_iStartIndex = tempString.indexOf(":") + 1;
						s_jsonOutputPath = tempString.substring(l_iStartIndex).trim();
					}
					// 是否需要主键
					else if (tempString.startsWith("WithMainKey")) {
						if (tempString.contains("yes") || tempString.contains("YES")) {
							s_bIsWithMainKey = true;
						} else {
							s_bIsWithMainKey = false;
						}
					}
					// 是否需要加密
					else if (tempString.startsWith("IS_Encrypt")) {
						if (tempString.contains("yes") || tempString.contains("YES")) {
							s_bIsEncrypt = true;
						} else {
							s_bIsEncrypt = false;
						}
					}
					// 产生校验码的输出替换文件
					else if (tempString.startsWith("MD5Code_File")) {
						int l_iStartIndex = tempString.indexOf(":") + 1;
						s_globalDataFileName = tempString.substring(l_iStartIndex).trim();
						see(s_globalDataFileName);
					}

					// >>>>>>>>>>>>>>>>>>>数据文件选项<<<<<<<<<<<<<<<<<<
					// Excel文件
					else if (tempString.startsWith("#")) {
						tempStructure = new XLSStructure(xlsInputPath + tempString.substring(1));
						xlsFileList.add(tempStructure);
					}
					// Sheet表
					else if (tempString.startsWith("-")) {
						if (tempStructure != null) {
							tempStructure.xlsSheets.add(tempString.substring(1));
						}
					}
					// 配置文件格式有错
					else {
						System.out.println(">>>>>>>>>>>>format error in conf.txt \"" + tempString + "\"<<<<<<<<<<<<");
					}
				}

				line++;
			}
			reader.close();
			analyzeConvert();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (reader != null) {
				try {
					reader.close();

				} catch (IOException e1) {
				}
			}

		}

	}

	public static void analyzeConvert() {
		System.out.println("\nAnalyzing...");
		for (int i = 0; i < xlsFileList.size(); i++) {
			//临时输出到本地目录下
			xlsFileList.get(i).AnalyzeJson(sf_strTempDir);
		}
	}

	//转换成json后的处理
	public static void processAfterJson(){
		//BufferedReader l_br;
		//BufferedReader l_subBr;
		BufferedWriter l_subBw;
		
		System.out.println("\n>>>>>>>>>>>>>processAfterJson");
		s_strMd5Array = new ArrayList<FileAndMd5Class>();
		try {
			// all excel
			for (int i = 0; i < xlsFileList.size(); i++) {
				int l_iSheetCount = xlsFileList.get(i).xlsSheets.size();
				// all sheet
				for (int j = 0; j < l_iSheetCount; j++) {
					String l_strSheetName = xlsFileList.get(i).xlsSheets.get(j);

					String l_strJsonFile_original = sf_strTempDir + l_strSheetName + ".txt";
					String l_strJsonFile_target = sf_strCurrentDir + l_strSheetName + ".txt";
					String l_strJsonFile_final = s_jsonOutputPath + l_strSheetName + ".txt";
					String l_stringContent;
					if (s_bIsEncrypt) {
						see(l_strSheetName + " encrypt by base64");
						l_stringContent = Base64Util.encode(AfterJsonTool.getByteArrayByFile(l_strJsonFile_original));
						
					}else{
						see(l_strSheetName + " without encrypt");
						l_stringContent = new String(AfterJsonTool.getByteArrayByFile(l_strJsonFile_original), "utf-8");
					}
					l_subBw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(l_strJsonFile_target), "utf-8"));
					l_subBw.write(l_stringContent);
					l_subBw.close();
										
					//产生校验码
					FileAndMd5Class l_fileAndMd5 = new FileAndMd5Class();
					l_fileAndMd5.fileName = l_strSheetName;
					l_fileAndMd5.md5CheckCode = AfterJsonTool.getMd5CodeByFileName(l_strJsonFile_target);
					s_strMd5Array.add(l_fileAndMd5);
															
					//移除文件到最终目录下
					File l_targetFile = new File(l_strJsonFile_target);
					File l_finalFile = new File(l_strJsonFile_final);
					l_targetFile.renameTo(l_finalFile);
				}
			}
			
			//把校验码写到指定文件中去
			see("\noutput md5checkCode to sepecial file....");
			AfterJsonTool.insertMd5CheckCode();
			see("All Process Over...\n\n");
						
		} catch (Exception ex) {
			System.out.println(">" + ex.toString());
			ex.printStackTrace();
		}
	}
	
	//输出调试信息
	public static void see(String p_string){
		System.out.println(p_string);
	}
}
