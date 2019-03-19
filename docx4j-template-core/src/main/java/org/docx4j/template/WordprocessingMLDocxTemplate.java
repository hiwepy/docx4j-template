/*
 * Copyright (c) 2018, vindell (https://github.com/vindell).
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not
 * use this file except in compliance with the License. You may obtain a copy of
 * the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
 * License for the specific language governing permissions and limitations under
 * the License.
 */
package org.docx4j.template;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.Docx4jProperties;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.template.utils.WmlZipUtils;

/**
 * 该模板负责对WordprocessingMLPackage进行普通变量替换和复杂变量替换并返回处理后的WordprocessingMLPackage对象
 * @author <a href="https://github.com/vindell">vindell</a>
 */
public class WordprocessingMLDocxTemplate extends WordprocessingMLTemplate {

	protected static final String PATH_TO_CONTENT = "word/document.xml";
	protected File sourceDocx;
	protected File outputDocx;
	protected File unzipDir;
	protected String placeholderStart = "${";
	protected String placeholderEnd = "}";
	protected String inputEncoding = Docx4jConstants.DEFAULT_CHARSETNAME;
	protected String outputEncoding = Docx4jConstants.DEFAULT_CHARSETNAME;
	protected boolean sourceDel = false;
	
	public WordprocessingMLDocxTemplate() {
	}
	
	public WordprocessingMLPackage process(String template, Map<String, Object> variables, File sourceDocx, File outputDocx) throws Exception {
		this.sourceDocx = sourceDocx;
		this.outputDocx = outputDocx;
		return this.process(template, variables);
	}
	
	/**
	 * 变量替换方式实现（只能解决固定模板的word生成）
	 * @param template ：模板内容
	 * @param variables ：变量
	 * @return {@link WordprocessingMLPackage} 对象
	 * @throws Exception ：异常对象
	 */
	@Override
	public WordprocessingMLPackage process(String template, Map<String, Object> variables) throws Exception {
		//初始化参数
		configuration();
		/*
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(sourceDocx);
		MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();  
        //获取静态变量集合
        HashMap<String, String> staticMap = getStaticData(variables);
        // 替换普通变量  
        documentPart.variableReplace(staticMap);  
        */
		//解压文档到指定的临时目录
		WmlZipUtils.unzip(this.sourceDocx, this.unzipDir);
		//加载主要文件
		File contentXmlFile = new File(this.unzipDir, PATH_TO_CONTENT);
		if (!contentXmlFile.exists()) {
			throw new FileNotFoundException(contentXmlFile.getAbsolutePath());
		}
		//进行变量替换
		//String template = FileUtils.readFileToString(contentXmlFile, this.inputEncoding );
		for (Map.Entry<String, Object> entry : variables.entrySet()) {
			//替换变量
			template = StringUtils.replace(template, this.placeholderStart + entry.getKey() + this.placeholderEnd, String.valueOf(entry.getValue()));
		}
		/*if(this.sourceDel){
			//删除源文件
			this.sourceDocx.delete();
		}*/
		//创建临时目录
		String sourceDocxPath = this.sourceDocx.getPath();
		String tempDocxPath = sourceDocxPath.substring(0, sourceDocxPath.lastIndexOf(".")) + "_bak";
		File tempDir = new File(tempDocxPath);
		//清空临时目录
		FileUtils.deleteDirectory(tempDir);
		//拷贝解压的文件目录到指定的文件目录
		FileUtils.copyDirectory(this.unzipDir, tempDir);
		//将处理后的主文档文件写到临时目录
		FileUtils.writeStringToFile(new File(tempDir, PATH_TO_CONTENT), template, this.outputEncoding );
		//把临时目录压缩为一个文件，且地址为源文件地址，以便替换掉源文件
		WmlZipUtils.zipDir(tempDir, this.outputDocx, false);
		//删除临时目录
		FileUtils.deleteDirectory(tempDir);
        //返回WordprocessingMLPackage对象
		return WordprocessingMLPackage.load(this.outputDocx);
	}
	
	protected void configuration() throws IOException{
		//初始化解压目录
		this.unzipDir = new File(this.sourceDocx.getParentFile(), Docx4jProperties.getProperty("docx4j.docx.tmpdir", "unzip_tmpdir"));
		if(!this.unzipDir.exists() || this.unzipDir.isFile()){
			this.unzipDir.setReadable(true);
			this.unzipDir.setWritable(true);
			this.unzipDir.mkdir();
		}
		this.inputEncoding = Docx4jProperties.getProperty("docx4j.docx.input.encoding", Docx4jConstants.DEFAULT_CHARSETNAME);
		this.outputEncoding = Docx4jProperties.getProperty("docx4j.docx.output.encoding", Docx4jConstants.DEFAULT_CHARSETNAME);
		this.placeholderStart = Docx4jProperties.getProperty("docx4j.docx.placeholderStart", "${");
		this.placeholderEnd = Docx4jProperties.getProperty("docx4j.docx.placeholderEnd", "}");
		this.sourceDel = Docx4jProperties.getProperty("docx4j.docx.source.delete", false);
	}
	
	/** 
     * 获取静态数据
     */
	protected HashMap<String, String> getStaticData(Map<String, Object> variables) { 
    	//静态数据集合
        HashMap<String, String> dataMap = new HashMap<String, String>();  
        if (variables != null) {
			for (String key : variables.keySet()) {
				Object val = variables.get(key);
				dataMap.put(key, val == null ? "" : val.toString()); 
			}
		}
        return dataMap;  
    }  

}
