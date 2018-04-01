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
package org.docx4j.template.rythm;

import java.io.IOException;
import java.io.StringWriter;
import java.net.URISyntaxException;
import java.util.Map;
import java.util.Properties;

import org.docx4j.Docx4jProperties;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.template.WordprocessingMLHtmlTemplate;
import org.docx4j.template.WordprocessingMLTemplate;
import org.docx4j.template.utils.ConfigUtils;
import org.rythmengine.Rythm;
import org.rythmengine.RythmEngine;

/**
 * 该模板仅负责使用Rythm模板引擎将指定模板生成HTML并将HTML转换成XHTML后，作为模板生成WordprocessingMLPackage对象
 * @author <a href="https://github.com/vindell">vindell</a>
 */
public class WordprocessingMLRythmTemplate extends WordprocessingMLTemplate {
	
	protected RythmEngine engine;
	protected WordprocessingMLHtmlTemplate mlHtmlTemplate;

	public WordprocessingMLRythmTemplate(boolean altChunk) {
		this.mlHtmlTemplate = new WordprocessingMLHtmlTemplate(altChunk) ;
	}
	
	public WordprocessingMLRythmTemplate(WordprocessingMLHtmlTemplate template) {
		this.mlHtmlTemplate = template;
	}

	@Override
	public WordprocessingMLPackage process(String template, Map<String, Object> variables) throws Exception {
		// 创建模板输出内容接收对象
		StringWriter output = new StringWriter();
		// 使用Rythm模板引擎渲染模板
		getEngine().getTemplate(template , variables).render(output);
		//获取模板渲染后的结果
		String html = output.toString();
		//使用HtmlTemplate进行渲染
		return mlHtmlTemplate.process(html, variables);
	}
	
	public RythmEngine getEngine() throws IOException {
		return engine == null ? getInternalEngine() : engine;
	}

	public void setEngine(RythmEngine engine) {
		this.engine = engine;
	}
	
	protected RythmEngine getInternalEngine() throws IOException{
		Properties props =  ConfigUtils.filterWithPrefix("docx4j.rythm.", "docx4j.rythm.", Docx4jProperties.getProperties(), false);
		
		props.put("engine.mode", Rythm.Mode.valueOf(props.getProperty("engine.mode", "dev")));
		props.put("log.enabled", false);
		props.put("feature.smart_escape.enabled", false);
		props.put("feature.transform.enabled", false);
		try {
			props.put("home.template", Rythm.class.getResource(props.getProperty("home.template")).toURI().toURL().getFile());
		} catch (URISyntaxException e) {
			// ignore
			props.put("home.tmp", "/");
		}
		props.put("codegen.dynamic_exp.enabled", true);
		props.put("built_in.code_type", "false");
		props.put("built_in.transformer", "false");
		props.put("engine.file_write", "false");
		props.put("codegen.compact.enabled", "false");
        
        return new RythmEngine(props);
	}

}
