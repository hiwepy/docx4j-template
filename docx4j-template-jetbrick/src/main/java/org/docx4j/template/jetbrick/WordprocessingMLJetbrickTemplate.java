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
package org.docx4j.template.jetbrick;

import java.io.IOException;
import java.io.StringWriter;
import java.util.Map;
import java.util.Properties;

import org.docx4j.Docx4jProperties;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.template.WordprocessingMLHtmlTemplate;
import org.docx4j.template.WordprocessingMLTemplate;
import org.docx4j.template.utils.ConfigUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import jetbrick.config.ConfigLoader;
import jetbrick.template.JetConfig;
import jetbrick.template.JetEngine;

/**
 * @className	： WordprocessingMLJetbrickTemplate
 * @description	： 该模板仅负责使用Jetbrick模板引擎将指定模板生成HTML并将HTML转换成XHTML后，作为模板生成WordprocessingMLPackage对象
 * @author 		： <a href="https://github.com/vindell">vindell</a>
 * @date		： 2017年5月24日 下午10:30:22
 * @version 	V1.0
 */
public class WordprocessingMLJetbrickTemplate extends WordprocessingMLTemplate {
	
	protected final Logger LOG = LoggerFactory.getLogger(WordprocessingMLJetbrickTemplate.class);
	protected JetEngine engine;
	protected WordprocessingMLHtmlTemplate mlHtmlTemplate;

	public WordprocessingMLJetbrickTemplate(boolean altChunk) {
		this.mlHtmlTemplate = new WordprocessingMLHtmlTemplate(altChunk) ;
	}
	
	public WordprocessingMLJetbrickTemplate(WordprocessingMLHtmlTemplate template) {
		this.mlHtmlTemplate = template;
	}

	@Override
	public WordprocessingMLPackage process(String template, Map<String, Object> variables) throws Exception {
		// 创建模板输出内容接收对象
		StringWriter output = new StringWriter();
		// 使用Httl模板引擎渲染模板
		getEngine().getTemplate(template).render(variables, output);
		//获取模板渲染后的结果
		String html = output.toString();
		//使用HtmlTemplate进行渲染
		return mlHtmlTemplate.process(html, variables);
	}
	
	public JetEngine getEngine() throws IOException {
		return engine == null ? getInternalEngine() : engine;
	}

	public void setEngine(JetEngine engine) {
		this.engine = engine;
	}
	
	protected JetEngine getInternalEngine() throws IOException{
		Properties ps = new Properties();
		ConfigLoader loader = new ConfigLoader();
		try {
			LOG.info("Loading config file: {}", JetConfig.DEFAULT_CONFIG_FILE);
		    loader.load(JetConfig.DEFAULT_CONFIG_FILE);
		    ps = loader.asProperties();
		} catch (Exception e) {
		     // 默认配置文件不存在
			LOG.warn("No default config file found: {}", JetConfig.DEFAULT_CONFIG_FILE);
			ps = ConfigUtils.filterWithPrefix("docx4j.jetx.", "docx4j.", Docx4jProperties.getProperties(), true);
		}
         return JetEngine.create(ps);
	}

}
