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
import java.util.HashMap;
import java.util.Map;

import org.docx4j.Docx4jProperties;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import webit.script.CFG;
import webit.script.Engine;

/**
 * @className	： WordprocessingMLWebitTemplate
 * @description	： 该模板仅负责使用Webit模板引擎将指定模板生成HTML并将HTML转换成XHTML后，作为模板生成WordprocessingMLPackage对象
 * @author 		： <a href="https://github.com/vindell">vindell</a>
 * @date		： 2017年5月24日 下午10:31:48
 * @version 	V1.0
 */
public class WordprocessingMLWebitTemplate extends WordprocessingMLTemplate {
	
	protected Engine engine;
	protected String templateKey;
	protected boolean altChunk = false ;
	
	public WordprocessingMLWebitTemplate(String template,boolean altChunk) {
		this.templateKey = template;
		this.altChunk = altChunk;
	}

	@Override
	public WordprocessingMLPackage process(Map<String, Object> variables) throws Exception {
		//创建模板输出内容接收对象
		StringWriter output = new StringWriter();
		//使用Webit模板引擎渲染模板
		getEngine().getTemplate(templateKey).merge(variables,output);
		//获取模板渲染后的结果
		String html = output.toString();
		//使用HtmlTemplate进行渲染
		return new WordprocessingMLHtmlTemplate(html , altChunk).process(variables);
	}
	
	public Engine getEngine() throws IOException {
		return engine == null ? getInternalEngine() : engine;
	}

	public void setEngine(Engine engine) {
		this.engine = engine;
	}

	protected Engine getInternalEngine() throws IOException{
		
		Map<String, Object> ps = new HashMap<String, Object>();
		ps.put(CFG.APPEND_LOST_SUFFIX, Docx4jProperties.getProperty("docx4j.webit.engine.appendLostSuffix", false));
		ps.put(CFG.INIT_TEMPLATES, Docx4jProperties.getProperty("docx4j.webit.engine.initTemplates"));
		//ps.put(CFG.FILTER, input_charset);
		ps.put(CFG.LOADER, Docx4jProperties.getProperty("docx4j.webit.engine.resourceLoader","webit.script.loaders.impl.ClasspathLoader"));
        ps.put(CFG.LOADER_ENCODING, Docx4jProperties.getProperty("docx4j.webit.loader.encoding", Engine.UTF_8) );
        ps.put(CFG.LOADER_ROOT, Docx4jProperties.getProperty("docx4j.webit.loader.root") );
        ps.put(CFG.LOGGER, Docx4jProperties.getProperty("docx4j.webit.engine.logger", "webit.script.loggers.impl.NOPLogger"));
		ps.put(CFG.LOOSE_VAR, Docx4jProperties.getProperty("docx4j.webit.engine.looseVar", false));
		ps.put(CFG.OUT_ENCODING, Docx4jProperties.getProperty("docx4j.webit.engine.encoding", Engine.UTF_8));
		//ps.put(CFG.RESOLVERS, input_charset);
		//ps.put(CFG.SECURITY_LIST, Docx4jProperties.getProperty("docx4j.webit.nativeSecurity.list"));
        //ps.put(CFG.SERVLET_CONTEXT, output_charset );
        ps.put(CFG.SHARE_ROOT, Docx4jProperties.getProperty("docx4j.webit.engine.shareRootData", true));
        ps.put(CFG.SUFFIX, Docx4jProperties.getProperty("docx4j.webit.engine.suffix", ".wit"));
        ps.put(CFG.TEXT_FACTORY, Docx4jProperties.getProperty("docx4j.webit.engine.textStatementFactory", CFG.SIMPLE_TEXT_FACTORY));
        ps.put(CFG.TRIM_CODE_LINE, Docx4jProperties.getProperty("docx4j.webit.engine.trimCodeBlockBlankLine",true));
        ps.put(CFG.VARS, Docx4jProperties.getProperty("docx4j.webit.engine.vars"));
        
        return Engine.create("", ps);
	}
}
