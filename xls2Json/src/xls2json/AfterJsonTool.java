package xls2json;

import xls2json.MD5Util;
import xls2json.Entrance;

import java.io.BufferedInputStream;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;

import xls2json.FileAndMd5Class;

public class AfterJsonTool {
	//根据文件产生md5码
		public static String getMd5CodeByFileName(String p_strFileName){
			
			try{
				return MD5Util.getMD5(getByteArrayByFile(p_strFileName));
			}catch (Exception ex){
				see(">"+ex.toString());
				ex.printStackTrace();
			}
			return "";
		}
		
		//产生新的GlobalData文件
		public static String insertMd5CheckCode(){
			BufferedReader l_br;
			BufferedWriter l_bw;
			//除路径之外的文件名 eg： "GlobalData.cpp"
			String l_targetGlobalFileName;
			int l_fileNameLabelIndex = Entrance.s_globalDataFileName.lastIndexOf("/");
			l_targetGlobalFileName = Entrance.s_globalDataFileName.substring(l_fileNameLabelIndex+1);
			try{
				l_br = new BufferedReader(new InputStreamReader(new FileInputStream(Entrance.s_globalDataFileName), "utf-8"));
				l_bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(l_targetGlobalFileName), "utf-8"));
				
				String l_strLine  = l_br.readLine();
				String l_strLineTarget;
				while(l_strLine != null){
					
					
					//插入新的宏定义
					if((l_strLine.indexOf(Entrance.sf_strInserterLabel) != -1) && Entrance.s_bIsMd5Check){
						l_strLineTarget = l_strLine;
						l_bw.write(l_strLineTarget + "\n");
						
						//插入新的宏
						for(int i=0; i<Entrance.s_strMd5Array.size(); i++){
							FileAndMd5Class l_fileAndMd5Class = Entrance.s_strMd5Array.get(i);
							l_bw.write("#define MD5_" + l_fileAndMd5Class.fileName + " \""+ l_fileAndMd5Class.md5CheckCode + "\"\n");						
						}
						
						//产生的文件是否加密，以宏的形式写到代码中
						if(Entrance.s_bIsEncrypt){
							l_bw.write("#define IsJsonFileEncrypted true\n");
						}else{
							l_bw.write("#define IsJsonFileEncrypted false\n");
						}					
					}
					
					//去掉原来的宏定义
					else if(l_strLine.indexOf("#define MD5_") != -1){
						//nothing
					}else if(l_strLine.indexOf("#define IsJsonFileEncrypted") != -1){
						//nothing
					}else{
						l_strLineTarget = l_strLine;
						l_bw.write(l_strLineTarget + "\n");
					}
					
					l_strLine = l_br.readLine();
				}
				l_br.close();
				l_bw.close();
				
				//移除到final文件
				File l_targetFile = new File(l_targetGlobalFileName);
				File l_finalFile = new File(Entrance.s_globalDataFileName);
				l_targetFile.renameTo(l_finalFile);
				
				
			}catch(Exception ex){
				see(">"+ex.toString());
				ex.printStackTrace();
			}
			return "";
		}
		
		//把一个文件转生byte[]
		public static byte[] getByteArrayByFile(String filename) throws IOException{  
	        
	        File f = new File(filename);  
	        if(!f.exists()){  
	            throw new FileNotFoundException(filename);  
	        }  
	  
	        ByteArrayOutputStream bos = new ByteArrayOutputStream((int)f.length());  
	        BufferedInputStream in = null;  
	        try{  
	            in = new BufferedInputStream(new FileInputStream(f));  
	            int buf_size = 1024;  
	            byte[] buffer = new byte[buf_size];  
	            int len = 0;  
	            while(-1 != (len = in.read(buffer,0,buf_size))){  
	                bos.write(buffer,0,len);  
	            }  
	            return bos.toByteArray();  
	        }catch (IOException e) {  
	            e.printStackTrace();  
	            throw e;  
	        }finally{  
	            try{  
	                in.close();  
	            }catch (IOException e) {  
	                e.printStackTrace();  
	            }  
	            bos.close();  
	        }  
	    }
		

		
		//输出调试信息
		public static void see(String p_string){
			System.out.println(p_string);
		}
}
