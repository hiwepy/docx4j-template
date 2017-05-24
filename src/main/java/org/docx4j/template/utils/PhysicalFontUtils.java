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
package org.docx4j.template.utils;

import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.template.fonts.ChineseFont;
import org.docx4j.template.fonts.PhysicalFontFactory;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;

/**
 * @className	： PhysicalFontUtils
 * @description	： TODO(描述这个类的作用)
 * @author 		： <a href="https://github.com/vindell">vindell</a>
 * @date		： 2017年5月24日 下午10:27:37
 * @version 	V1.0
 */
public class PhysicalFontUtils {

	
	private static PhysicalFontFactory FONT_FACTORY = PhysicalFontFactory.getInstance();
 
	/**
	 * 为 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage} 配置中文字体;解决中文乱码问题
	 */
	public static void setWmlPackageFonts(WordprocessingMLPackage wmlPackage) throws Docx4JException {
		try {
			//字体映射;  
			Mapper fontMapper = FONT_FACTORY.getFontMapper();;
			//设置文档字体库
			wmlPackage.setFontMapper(fontMapper, true);
		} catch (Exception e) {
			throw new Docx4JException(e.getMessage(),e.getCause());
		}
    }
	
	/**
	 * 为 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage} 配置默认字体
	 */
	public static void setDefaultFont(WordprocessingMLPackage wmlPackage,String fontName) throws Docx4JException {
        //设置文件默认字体
        RFonts rfonts = Context.getWmlObjectFactory().createRFonts(); 
        rfonts.setAsciiTheme(null);
        rfonts.setAscii(fontName);
        rfonts.setHAnsi(fontName);
        rfonts.setEastAsia(fontName);
        RPr rpr = wmlPackage.getMainDocumentPart().getPropertyResolver().getDocumentDefaultRPr();
        rpr.setRFonts(rfonts);
    }
	
	/**
	 * 为 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage} 配置中文字体
	 */
	public static void setSimSunFont(WordprocessingMLPackage wmlPackage) throws Docx4JException {
        //设置文件默认字体
		setDefaultFont(wmlPackage,ChineseFont.SIMSUM.getFontName());
    }
	
	/**
	 * 为 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage} 增加新的字体
	 */
	public static void setPhysicalFont(WordprocessingMLPackage wmlPackage,PhysicalFont physicalFont) throws Exception {
		Mapper fontMapper = wmlPackage.getFontMapper() == null ? new IdentityPlusMapper() : wmlPackage.getFontMapper();
		//分别设置字体名和别名对应的字体库
		fontMapper.put(physicalFont.getName(), physicalFont );
		//设置文档字体库
		wmlPackage.setFontMapper(fontMapper, true);
    }
	
	/**
	 * 为 {@link org.docx4j.openpackaging.packages.WordprocessingMLPackage} 增加新的字体
	 */
	public static void setPhysicalFont(WordprocessingMLPackage wmlPackage,String fontName) throws Exception {
		//Mapper fontMapper = new BestMatchingMapper();  
		Mapper fontMapper = wmlPackage.getFontMapper() == null ? new IdentityPlusMapper() : wmlPackage.getFontMapper();
		//获取字体库
		PhysicalFont physicalFont = PhysicalFonts.get(fontName);
		//分别设置字体名和别名对应的字体库
		fontMapper.put(fontName, physicalFont );
		//设置文档字体库
		wmlPackage.setFontMapper(fontMapper, true);
    }
}
