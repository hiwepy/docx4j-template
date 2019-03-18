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
package org.docx4j.template.xhtml.io;

import java.io.File;
import java.io.IOException;
import java.net.URL;

import org.docx4j.events.StartEvent;
import org.docx4j.model.structure.PageSizePaper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.template.events.BuildFinishedEvent;
import org.docx4j.template.events.BuildJobTypes;
import org.docx4j.template.events.BuildStartEvent;
import org.docx4j.template.fonts.ChineseFont;
import org.docx4j.template.utils.PhysicalFontUtils;
import org.docx4j.template.xhtml.DataMap;
import org.docx4j.template.xhtml.handler.DocumentHandler;
import org.docx4j.template.xhtml.handler.def.XHTMLDocumentHandler;
import org.docx4j.template.xhtml.utils.XHTMLImporterUtils;
import org.jsoup.nodes.Document;

public class WordprocessingMLPackageBuilder {

	protected DocumentHandler docHandler = new XHTMLDocumentHandler();
	
	/*
	 * 为 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage} 配置中文字体;解决中文乱码问题
	 */
	public WordprocessingMLPackageBuilder configChineseFonts(WordprocessingMLPackage wmlPackage) throws IOException, Docx4JException {
		//初始化中文字体
		PhysicalFontUtils.setWmlPackageFonts(wmlPackage);
        //返回WordprocessingMLPackage对象
      	return this;
    }
	
	/*
	 * 为 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage} 配置默认字体
	 */
	public WordprocessingMLPackageBuilder configDefaultFont(WordprocessingMLPackage wmlPackage,String fontName) throws IOException, Docx4JException {
		//设置文件默认字体
		try {
			PhysicalFontUtils.setDefaultFont(wmlPackage, fontName);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        //返回WordprocessingMLPackage对象
      	return this;
    }
	
	/*
	 * 为 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage} 配置中文字体
	 */
	public WordprocessingMLPackageBuilder configSimSunFont(WordprocessingMLPackage wmlPackage) throws IOException, Docx4JException {
		//初始化中文字体，解决中文乱码问题
		configChineseFonts(wmlPackage);
        //设置文件默认字体
		configDefaultFont(wmlPackage,ChineseFont.SIMSUM.getFontName());
		//返回WordprocessingMLPackage对象
		return this;
    }
	
	/*
	 * 获取初始化后的 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage}对象
	 * @param wmlPackage
	 * @return
	 */
	public WordprocessingMLPackage initialize(WordprocessingMLPackage wmlPackage) {
		
		/*MBassador<Docx4jEvent> bus = new MBassador<Docx4jEvent>();
			Docx4jEvent.setEventNotifier(bus);
		*/
		
		
		return wmlPackage;
	}
	
	public WordprocessingMLPackage buildWhithDoc(Document doc,boolean altChunk) throws IOException, Docx4JException {
		/*	
		 * 	根据docx4j.properties配置文件中:
		 * 	docx4j.PageSize = A4
		 * 	docx4j.PageOrientationLandscape = true
		 * 	创建默认的WordProcessingML package
		 */
        WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.createPackage();
        //返回WordprocessingMLPackage对象
        return buildWhithDoc(wmlPackage , doc, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithDoc(Document doc,boolean landscape,boolean altChunk) throws IOException, Docx4JException {
        //返回WordprocessingMLPackage对象
        return buildWhithDoc(doc, PageSizePaper.A4, landscape, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithDoc(Document doc,PageSizePaper pageSize,boolean landscape,boolean altChunk) throws IOException, Docx4JException {
		//创建指定纸张大小和方向的WordprocessingMLPackage对象
        WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.createPackage(pageSize, landscape); //A4纸，
        //返回WordprocessingMLPackage对象
        return buildWhithDoc(wmlPackage , doc , altChunk);
    }
	
	public WordprocessingMLPackage buildWhithDoc(WordprocessingMLPackage wmlPackage,Document doc,boolean altChunk) throws IOException, Docx4JException {
		//构建任务开始
		StartEvent jobStartEvent = new BuildStartEvent(BuildJobTypes.DOC, wmlPackage);
		jobStartEvent.publish();
		//配置中文字体
        wmlPackage = initialize(wmlPackage);
        //渲染WordprocessingMLPackage对象
  		XHTMLImporterUtils.handle(wmlPackage, doc, false, altChunk);
        //构建任务结束
        new BuildFinishedEvent(jobStartEvent).publish();
        //返回WordprocessingMLPackage对象
        return wmlPackage;
    }
	
	/*
	 * 将 {@link org.jsoup.nodes.Document} 对象转为 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage}
	 */
	public WordprocessingMLPackage buildWhithXhtml(File htmlFile,boolean altChunk) throws IOException, Docx4JException {
		/*	
		 * 	根据docx4j.properties配置文件中:
		 * 	docx4j.PageSize = A4
		 * 	docx4j.PageOrientationLandscape = true
		 * 	创建默认的WordProcessingML package
		 */
		WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.createPackage();
		//返回WordprocessingMLPackage对象
		return buildWhithXhtml(wmlPackage,htmlFile, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithXhtml(File htmlFile,boolean landscape,boolean altChunk) throws IOException, Docx4JException {
        //返回WordprocessingMLPackage对象
        return buildWhithXhtml(htmlFile, PageSizePaper.A4, landscape, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithXhtml(File htmlFile,PageSizePaper pageSize,boolean landscape ,boolean altChunk) throws IOException, Docx4JException {
		//创建指定纸张大小和方向的WordprocessingMLPackage对象
        WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.createPackage(pageSize, landscape); //A4纸，
        //返回WordprocessingMLPackage对象
        return buildWhithXhtml(wmlPackage, htmlFile , altChunk);
    }
	
	public WordprocessingMLPackage buildWhithXhtml(WordprocessingMLPackage wmlPackage,File htmlFile, boolean altChunk) throws IOException, Docx4JException{
		//构建任务开始
  		StartEvent jobStartEvent = new BuildStartEvent(BuildJobTypes.HTML, wmlPackage);
  		jobStartEvent.publish();
		//配置中文字体
		wmlPackage = initialize(wmlPackage); 
		//渲染WordprocessingMLPackage对象
  		XHTMLImporterUtils.handle(wmlPackage, docHandler.handle(htmlFile), false, altChunk);
		//构建任务结束
		new BuildFinishedEvent(jobStartEvent).publish();
		//返回WordprocessingMLPackage对象
		return wmlPackage;
    }
	
	/*
	 * 将 {@link org.jsoup.nodes.Document} 对象转为 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage}
	 */
	public WordprocessingMLPackage buildWhithXhtml(String html,boolean altChunk) throws IOException, Docx4JException {
		/*	
		 * 	根据docx4j.properties配置文件中:
		 * 	docx4j.PageSize = A4
		 * 	docx4j.PageOrientationLandscape = true
		 * 	创建默认的WordProcessingML package
		 */
  		WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.createPackage();
		//返回WordprocessingMLPackage对象
		return buildWhithXhtml(wmlPackage,html, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithXhtml(String html,boolean landscape,boolean altChunk) throws IOException, Docx4JException {
        //返回WordprocessingMLPackage对象
        return buildWhithXhtml(html, PageSizePaper.A4, landscape, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithXhtml(String html,PageSizePaper pageSize,boolean landscape,boolean altChunk) throws IOException, Docx4JException {
		//创建指定纸张大小和方向的WordprocessingMLPackage对象
        WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.createPackage(pageSize, landscape); //A4纸，
        //返回WordprocessingMLPackage对象
        return buildWhithXhtml(wmlPackage, html, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithXhtml(WordprocessingMLPackage wmlPackage,String html,boolean altChunk) throws IOException, Docx4JException {
		//构建任务开始
  		StartEvent jobStartEvent = new BuildStartEvent(BuildJobTypes.HTML, wmlPackage);
  		jobStartEvent.publish();
		//配置中文字体
		wmlPackage = initialize(wmlPackage); 
		//渲染WordprocessingMLPackage对象
  		XHTMLImporterUtils.handle(wmlPackage, docHandler.handle(html , false), false, altChunk);
		//构建任务结束
		new BuildFinishedEvent(jobStartEvent).publish();
		//返回WordprocessingMLPackage对象
		return wmlPackage;
    }
	
	/*
	 * 将 {@link org.jsoup.nodes.Document} 对象转为 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage}
	 */
	public WordprocessingMLPackage buildWhithXhtmlFragment(String xhtml,boolean altChunk) throws IOException, Docx4JException {
		/*	
		 * 	根据docx4j.properties配置文件中:
		 * 	docx4j.PageSize = A4
		 * 	docx4j.PageOrientationLandscape = true
		 * 	创建默认的WordProcessingML package
		 */
  		WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.createPackage();
		//返回WordprocessingMLPackage对象
		return buildWhithXhtmlFragment(wmlPackage, xhtml, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithXhtmlFragment(String html,boolean landscape, boolean altChunk) throws IOException, Docx4JException {
        //返回WordprocessingMLPackage对象
        return buildWhithXhtmlFragment(html, PageSizePaper.A4, landscape, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithXhtmlFragment(String xhtml,PageSizePaper pageSize,boolean landscape,boolean altChunk) throws IOException, Docx4JException {
		//创建指定纸张大小和方向的WordprocessingMLPackage对象
        WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.createPackage(pageSize, landscape); //A4纸，
        //返回WordprocessingMLPackage对象
        return buildWhithXhtmlFragment(wmlPackage, xhtml, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithXhtmlFragment(WordprocessingMLPackage wmlPackage,String xhtml, boolean altChunk) throws IOException, Docx4JException {
		//构建任务开始
  		StartEvent jobStartEvent = new BuildStartEvent(BuildJobTypes.HTML, wmlPackage);
  		jobStartEvent.publish();
		//配置中文字体
		wmlPackage = initialize(wmlPackage); 
		//渲染WordprocessingMLPackage对象
  		XHTMLImporterUtils.handle(wmlPackage, docHandler.handle(xhtml , true), true, altChunk);
		//构建任务结束
		new BuildFinishedEvent(jobStartEvent).publish();
		//返回WordprocessingMLPackage对象
		return wmlPackage;
    }
	
	/*
	 * 将 {@link org.jsoup.nodes.Document} 对象转为 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage}
	 */
	public WordprocessingMLPackage buildWhithURL(URL url, boolean altChunk) throws IOException, Docx4JException {
		/*	
		 * 	根据docx4j.properties配置文件中:
		 * 	docx4j.PageSize = A4
		 * 	docx4j.PageOrientationLandscape = true
		 * 	创建默认的WordProcessingML package
		 */
  		WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.createPackage();
		//返回WordprocessingMLPackage对象
		return buildWhithURL(wmlPackage,url, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithURL(URL url,boolean landscape, boolean altChunk) throws IOException, Docx4JException {
        //返回WordprocessingMLPackage对象
        return buildWhithURL(url, PageSizePaper.A4, landscape, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithURL(URL url,PageSizePaper pageSize,boolean landscape, boolean altChunk) throws IOException, Docx4JException {
		//创建指定纸张大小和方向的WordprocessingMLPackage对象
        WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.createPackage(pageSize, landscape); //A4纸，
        //返回WordprocessingMLPackage对象
        return buildWhithURL(wmlPackage, url, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithURL(WordprocessingMLPackage wmlPackage,URL url, boolean altChunk) throws IOException, Docx4JException {
		//构建任务开始
  		StartEvent jobStartEvent = new BuildStartEvent(BuildJobTypes.URL, wmlPackage);
  		jobStartEvent.publish();
		//配置中文字体
		wmlPackage = initialize(wmlPackage); 
		//渲染WordprocessingMLPackage对象
  		XHTMLImporterUtils.handle(wmlPackage, docHandler.handle(url), false, altChunk);
		//构建任务结束
		new BuildFinishedEvent(jobStartEvent).publish();
		//返回WordprocessingMLPackage对象
		return wmlPackage;
    }
	
	public WordprocessingMLPackage buildWhithURL(String url, DataMap dataMap, boolean altChunk) throws IOException, Docx4JException {
		/*	
		 * 	根据docx4j.properties配置文件中:
		 * 	docx4j.PageSize = A4
		 * 	docx4j.PageOrientationLandscape = true
		 * 	创建默认的WordProcessingML package
		 */
  		WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.createPackage();
		//返回WordprocessingMLPackage对象
		return buildWhithURL(wmlPackage, url, dataMap, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithURL(String url, DataMap dataMap,boolean landscape, boolean altChunk) throws IOException, Docx4JException {
        //返回WordprocessingMLPackage对象
        return buildWhithURL(url, dataMap, PageSizePaper.A4, landscape, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithURL(String url, DataMap dataMap,PageSizePaper pageSize,boolean landscape, boolean altChunk) throws IOException, Docx4JException {
		//创建指定纸张大小和方向的WordprocessingMLPackage对象
        WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.createPackage(pageSize, landscape); //A4纸，
        //返回WordprocessingMLPackage对象
        return buildWhithURL(wmlPackage, url, dataMap, altChunk);
    }
	
	public WordprocessingMLPackage buildWhithURL(WordprocessingMLPackage wmlPackage,String url, DataMap dataMap, boolean altChunk) throws IOException, Docx4JException {
		//构建任务开始
  		StartEvent jobStartEvent = new BuildStartEvent(BuildJobTypes.URL, wmlPackage);
  		jobStartEvent.publish();
		//配置中文字体
		wmlPackage = initialize(wmlPackage); 
		//渲染WordprocessingMLPackage对象
  		XHTMLImporterUtils.handle(wmlPackage, docHandler.handle(url, dataMap), false, altChunk);
		//构建任务结束
		new BuildFinishedEvent(jobStartEvent).publish();
		//返回WordprocessingMLPackage对象
		return wmlPackage;
    }
	
}
