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
package org.docx4j.template.xhtml;

import java.io.File;
import java.io.InputStream;
import java.net.URL;
import java.util.Map;

import org.apache.commons.io.IOUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.template.WordprocessingMLTemplate;
import org.docx4j.template.xhtml.handler.DocumentHandler;
import org.docx4j.template.xhtml.handler.def.XHTMLDocumentHandler;
import org.docx4j.template.xhtml.io.WordprocessingMLPackageBuilder;
import org.jsoup.nodes.Document;

/**
 * 该模板仅负责将原生的HTML元素转换成XHTML后，作为模板生成WordprocessingMLPackage对象
 * @author <a href="https://github.com/vindell">vindell</a>
 */
public class WordprocessingMLHtmlTemplate extends WordprocessingMLTemplate {

	protected DocumentHandler docHandler = new XHTMLDocumentHandler();
	protected WordprocessingMLPackageBuilder wordMLPackageBuilder = new WordprocessingMLPackageBuilder();
	protected File htmlFile;
	protected URL url;
	protected String urlstr;
	protected DataMap dataMap;
	protected InputStream input;
	protected Document doc;
	protected boolean altChunk;

	public WordprocessingMLHtmlTemplate(File htmlFile, boolean altChunk) {
		this.htmlFile = htmlFile;
		this.altChunk = altChunk;
	}

	public WordprocessingMLHtmlTemplate(boolean altChunk) {
		this.altChunk = altChunk;
	}

	public WordprocessingMLHtmlTemplate(URL url, boolean altChunk) {
		this.url = url;
		this.altChunk = altChunk;
	}

	public WordprocessingMLHtmlTemplate(String url, DataMap dataMap, boolean altChunk) {
		this.urlstr = url;
		this.dataMap = dataMap;
		this.altChunk = altChunk;
	}

	public WordprocessingMLHtmlTemplate(InputStream input, boolean altChunk) {
		this.input = input;
		this.altChunk = altChunk;
	}

	public WordprocessingMLHtmlTemplate(Document doc, boolean altChunk) {
		this.doc = doc;
		this.altChunk = altChunk;
	}

	/**
	 * 将 {@link org.jsoup.nodes.Document} 对象转为 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage}
	 * @param template ：模板内容
	 * @param variables ：变量
	 * @return {@link WordprocessingMLPackage} 对象
	 * @throws Exception ：异常对象
	 */
	@Override
	public WordprocessingMLPackage process(String template, Map<String, Object> variables) throws Exception {
		WordprocessingMLPackage wordMLPackage = null;
		if (htmlFile != null) {
			wordMLPackage = wordMLPackageBuilder.buildWhithXhtml(htmlFile, altChunk);
		} else if (url != null) {
			wordMLPackage = wordMLPackageBuilder.buildWhithURL(url, altChunk);
		} else if (urlstr != null) {
			wordMLPackage = wordMLPackageBuilder.buildWhithURL(urlstr, dataMap, altChunk);
		} else if (input != null) {
			try {
				wordMLPackage = wordMLPackageBuilder.buildWhithDoc(docHandler.handle(input), altChunk);
			} finally {
				IOUtils.closeQuietly(input);
			}
		} else if (doc != null) {
			wordMLPackage = wordMLPackageBuilder.buildWhithDoc(doc, altChunk);
		} else {
			wordMLPackage = wordMLPackageBuilder.buildWhithXhtml(template, altChunk);
		}
		// 返回WordprocessingMLPackage对象
		return wordMLPackage;
	}

	public DocumentHandler getDocHandler() {
		return docHandler;
	}

	public void setDocHandler(DocumentHandler docHandler) {
		this.docHandler = docHandler;
	}

	public WordprocessingMLPackageBuilder getWordMLPackageBuilder() {
		return wordMLPackageBuilder;
	}

	public void setWordMLPackageBuilder(WordprocessingMLPackageBuilder wordMLPackageBuilder) {
		this.wordMLPackageBuilder = wordMLPackageBuilder;
	}

	public File getHtmlFile() {
		return htmlFile;
	}

	public void setHtmlFile(File htmlFile) {
		this.htmlFile = htmlFile;
	}

	public URL getUrl() {
		return url;
	}

	public void setUrl(URL url) {
		this.url = url;
	}

	public String getUrlstr() {
		return urlstr;
	}

	public void setUrlstr(String urlstr) {
		this.urlstr = urlstr;
	}

	public DataMap getDataMap() {
		return dataMap;
	}

	public void setDataMap(DataMap dataMap) {
		this.dataMap = dataMap;
	}

	public InputStream getInput() {
		return input;
	}

	public void setInput(InputStream input) {
		this.input = input;
	}

	public Document getDoc() {
		return doc;
	}

	public void setDoc(Document doc) {
		this.doc = doc;
	}

	public boolean isAltChunk() {
		return altChunk;
	}

	public void setAltChunk(boolean altChunk) {
		this.altChunk = altChunk;
	}

}
