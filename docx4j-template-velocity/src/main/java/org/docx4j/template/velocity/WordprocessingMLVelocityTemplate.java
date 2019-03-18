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
package org.docx4j.template.velocity;

import java.io.IOException;
import java.io.StringWriter;
import java.util.Map;
import java.util.Properties;

import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.VelocityEngine;
import org.apache.velocity.tools.generic.DateTool;
import org.docx4j.Docx4jProperties;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.template.Docx4jConstants;
import org.docx4j.template.WordprocessingMLTemplate;
import org.docx4j.template.xhtml.WordprocessingMLHtmlTemplate;

/**
 * 该模板仅负责使用Velocity模板引擎将指定模板生成HTML并将HTML转换成XHTML后，作为模板生成WordprocessingMLPackage对象
 * @author <a href="https://github.com/vindell">vindell</a>
 */
public class WordprocessingMLVelocityTemplate extends WordprocessingMLTemplate {
	
	protected VelocityEngine engine;
	protected DateTool dateTool = new DateTool();
	protected WordprocessingMLHtmlTemplate mlHtmlTemplate;
	
	public WordprocessingMLVelocityTemplate(boolean altChunk) {
		this.mlHtmlTemplate = new WordprocessingMLHtmlTemplate(altChunk) ;
	}
	
	public WordprocessingMLVelocityTemplate(WordprocessingMLHtmlTemplate template) {
		this.mlHtmlTemplate = template;
	}
	
	/**
	 * 使用Velocity模板引擎渲染模板
	 * @param template ：模板内容
	 * @param variables ：变量
	 * @return {@link WordprocessingMLPackage} 对象
	 * @throws Exception ：异常对象
	 */
	@Override
	public WordprocessingMLPackage process(String template, Map<String, Object> variables) throws Exception {
		//设置Velocity上下文对象
		VelocityContext ctx = new VelocityContext(variables);
		ctx.put("dateTool", dateTool);
		// 创建模板输出内容接收对象
		StringWriter output = new StringWriter();
		// 使用Velocity模板引擎渲染模板
		getEngine().getTemplate(template).merge(ctx, output);
		//获取模板渲染后的结果
		String html = output.toString();
		//使用HtmlTemplate进行渲染
		return mlHtmlTemplate.process(html, variables);
	}
	
	public VelocityEngine getEngine() throws IOException {
		return engine == null ? getInternalEngine() : engine;
	}

	public void setEngine(VelocityEngine engine) {
		this.engine = engine;
	}

	protected VelocityEngine getInternalEngine() throws IOException{
		
		VelocityEngine engine = new VelocityEngine();
        
		Properties ps = new Properties();
        ps.setProperty(";runtime.log", Docx4jProperties.getProperty("docx4j.velocity.runtime.log", "velocity.log"));
        ps.setProperty(";runtime.log.logsystem.class", Docx4jProperties.getProperty("docx4j.velocity.runtime.log.logsystem.class", "org.apache.velocity.runtime.log.NullLogSystem"));
        ps.setProperty("resource.loader", Docx4jProperties.getProperty("docx4j.velocity.resource.loader", "file"));
        ps.setProperty("file.resource.loader.cache", Docx4jProperties.getProperty("docx4j.velocity.file.resource.loader.cache", "true"));
        ps.setProperty("file.resource.loader.class ", Docx4jProperties.getProperty("docx4j.velocity.file.resource.loader.class", "Velocity.Runtime.Resource.Loader.FileResourceLoader") );
        ps.setProperty(";resource.loader", Docx4jProperties.getProperty("docx4j.velocity.resource.loader", "webapp"));
        ps.setProperty(";webapp.resource.loader.class", Docx4jProperties.getProperty("docx4j.velocity.webapp.resource.loader.class", "org.apache.velocity.tools.view.servlet.WebappLoader"));
        ps.setProperty(";webapp.resource.loader.cache", Docx4jProperties.getProperty("docx4j.velocity.webapp.resource.loader.cache", "true"));
        ps.setProperty(";webapp.resource.loader.modificationCheckInterval", Docx4jProperties.getProperty("docx4j.velocity.webapp.resource.loader.modificationCheckInterval", "3") );
        ps.setProperty(";directive.foreach.counter.name", Docx4jProperties.getProperty("docx4j.velocity.directive.foreach.counter.name", "velocityCount"));
        ps.setProperty(";directive.foreach.counter.initial.value", Docx4jProperties.getProperty("docx4j.velocity.directive.foreach.counter.initial.value", "1"));
        ps.setProperty("file.resource.loader.path", this.getClass().getResource(Docx4jProperties.getProperty("docx4j.velocity.file.resource.loader.path", "/template")).getPath());
        //模板输入输出编码格式
        String input_charset = Docx4jProperties.getProperty("docx4j.velocity.input.encoding", Docx4jConstants.DEFAULT_CHARSETNAME);
        String output_charset = Docx4jProperties.getProperty("docx4j.velocity.output.encoding", Docx4jConstants.DEFAULT_CHARSETNAME );
        ps.setProperty("input.encoding", input_charset);
        ps.setProperty("output.encoding", output_charset);
        engine.init(ps);
        
        return engine;
	}

}
