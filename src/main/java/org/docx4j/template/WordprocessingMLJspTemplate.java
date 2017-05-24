/*
 * Copyright (c) 2010-2020, vindell (https://github.com/vindell).
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

import java.io.IOException;
import java.io.StringWriter;
import java.util.Map;
import java.util.Properties;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.docx4j.Docx4jProperties;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.template.engine.JspConfig;
import org.docx4j.template.engine.JspEngine;
import org.docx4j.template.utils.ConfigUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @className	： WordprocessingMLJspTemplate
 * @description	： 该模板仅负责使用Jsp模板引擎将指定模板生成HTML并将HTML转换成XHTML后，作为模板生成WordprocessingMLPackage对象
 * @author 		： <a href="https://github.com/vindell">vindell</a>
 * @date		： 2017年5月24日 下午10:30:34
 * @version 	V1.0
 */
public class WordprocessingMLJspTemplate extends WordprocessingMLTemplate {
	
    protected final Logger LOG = LoggerFactory.getLogger(WordprocessingMLJspTemplate.class);
	protected final HttpServletRequest request;
	protected final HttpServletResponse response;
	//要生成html的jsp文件路径（如：/frontStage/articleMenuContent.jsp）,这是实际存在的jsp文件
    protected final String name;
    protected final String requestURL;
    protected JspEngine engine;
	protected boolean altChunk = false ;
    
	public WordprocessingMLJspTemplate(HttpServletRequest request,HttpServletResponse response,String name, String requestURL,boolean altChunk) {
		this.request = request;
        this.response = response;
        this.name = name;
        this.requestURL = requestURL;
        this.altChunk = altChunk;
	}

	@Override
	public WordprocessingMLPackage process(Map<String, Object> variables) throws Exception {
		// 创建模板输出内容接收对象
		StringWriter output = new StringWriter();
		// 使用Jsp模板引擎渲染模板
		getEngine().getTemplate(request, response, name ).render(requestURL, variables, output);
		//获取模板渲染后的结果
		String html = output.toString();
		//使用HtmlTemplate进行渲染
		return new WordprocessingMLHtmlTemplate(html , altChunk).process(variables);
	}
	
	public JspEngine getEngine() throws IOException {
		return engine == null ? getInternalEngine() : engine;
	}

	public void setEngine(JspEngine engine) {
		this.engine = engine;
	}
	
	protected JspEngine getInternalEngine() throws IOException{
		Properties ps =  ConfigUtils.filterWithPrefix("docx4j.jsp.", "docx4j.jsp.", Docx4jProperties.getProperties(), false);
		//设置默认的参数
		ps.setProperty(JspConfig.TEMPLATE_SUFFIX, Docx4jProperties.getProperty("docx4j.jsp.template.suffix",".httl"));
		ps.setProperty(JspConfig.INPUT_ENCODING, Docx4jProperties.getProperty("docx4j.jsp.input.encoding", Docx4jConstants.DEFAULT_CHARSETNAME));
		ps.setProperty(JspConfig.OUTPUT_ENCODING, Docx4jProperties.getProperty("docx4j.jsp.output.encoding", Docx4jConstants.DEFAULT_CHARSETNAME));
        return JspEngine.create(ps);
	}

}
