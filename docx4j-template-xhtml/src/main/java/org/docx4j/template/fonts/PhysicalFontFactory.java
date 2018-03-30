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
package org.docx4j.template.fonts;

import java.net.MalformedURLException;
import java.net.URL;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;

import org.docx4j.Docx4jProperties;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.template.Docx4jConstants;
import org.docx4j.template.utils.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 
 * @className	： PhysicalFontFactory
 * @description	： TODO(描述这个类的作用)
 * @author 		： <a href="https://github.com/vindell">vindell</a>
 * @date		： 2017年5月24日 下午10:26:25
 * @version 	V1.0
 */
public class PhysicalFontFactory {

	protected static Logger LOG = LoggerFactory.getLogger(PhysicalFontFactory.class);
	protected static ConcurrentMap<String, FontScheme> COMPLIED_FONTSCHEME = new ConcurrentHashMap<String, FontScheme>();
	private volatile static PhysicalFontFactory singleton;

	public static PhysicalFontFactory getInstance() {
		if (singleton == null) {
			synchronized (PhysicalFontFactory.class) {
				if (singleton == null) {
					singleton = new PhysicalFontFactory();
				}
			}
		}
		return singleton;
	}
	
	private PhysicalFontFactory(){
		configExternalFonts();
		configInternalFonts();
	}
	
	public void configExternalFonts() {
		/**
		 * 加载外部字体库：参数格式如：字体名称:字体别名:字体URL；多个以 ",; \t\n"分割
		 * 隶书:LiSu:URL,宋体:SimSun:URL
		 */
		String external_fonts_mapping = Docx4jProperties.getProperty(Docx4jConstants.DOCX4J_FONTS_EXTERNAL_MAPPING);
		if(external_fonts_mapping!=null && external_fonts_mapping.length() > 0){
			String[] fonts_mappings = StringUtils.tokenizeToStringArray(external_fonts_mapping);
			if(fonts_mappings.length > 0){
				for (String fonts_mapping : fonts_mappings) {
					String[] font = fonts_mapping.split(":");
					FontScheme fontScheme = COMPLIED_FONTSCHEME.get(font[0]);
					if (fontScheme != null) {
						continue;
					}
					try {
						fontScheme = new FontScheme(font[0], font[1], new URL(font[2]));
						COMPLIED_FONTSCHEME.putIfAbsent(font[0], fontScheme);
						PhysicalFont physicalFont = PhysicalFonts.get(fontScheme.getFontName());
						if(physicalFont == null){
							//加载字体文件（解决linux环境下无中文字体问题）
							PhysicalFonts.addPhysicalFonts(fontScheme.getFontName(), fontScheme.getFontURL());
							LOG.debug("Add External " +  fontScheme.toString() );
						}
					} catch (MalformedURLException e) {
						// ignore
					}
				}
			}
		}
	}
	
	//加载字体文件（解决linux环境下无中文字体问题）
	public void configInternalFonts(){
        ChineseFont[] fonts = ChineseFont.values();
        for (ChineseFont font : fonts){
        	FontScheme fontScheme = COMPLIED_FONTSCHEME.get(font.getFontName());
			if (fontScheme == null) {
				try {
					fontScheme = new FontScheme(font.getFontName(), font.getFontAlias(), font.getFontURL());
					COMPLIED_FONTSCHEME.putIfAbsent(font.getFontName(), fontScheme);
					PhysicalFont physicalFont = PhysicalFonts.get(fontScheme.getFontName());
					if(physicalFont == null){
						//加载字体文件（解决linux环境下无中文字体问题）
						PhysicalFonts.addPhysicalFonts(fontScheme.getFontName(), fontScheme.getFontURL());
						LOG.debug("Add Internal " +  fontScheme.toString() );
					}
				} catch (Exception e) {
					LOG.debug("Add Internal PhysicalFont Fail : Ingore" );
				}
			}
        }
	}
	
	public Mapper getFontMapper() throws Exception {
		
		// Set up font mapper (optional)
		
		//  example of mapping font Times New Roman which doesn't have certain Arabic glyphs
		// eg Glyph "ي" (0x64a, afii57450) not available in font "TimesNewRomanPS-ItalicMT".
		// eg Glyph "ج" (0x62c, afii57420) not available in font "TimesNewRomanPS-ItalicMT".
		// to a font which does
		PhysicalFonts.get("Arial Unicode MS"); 
		
		
		/*
		 * This mapper uses Panose to guess the physical font which is a closest fit for the font used in the document.
		 * （这个映射器使用Panose算法猜测最适合这个文档使用的物理字体。）
		 * 
		 * Panose是一种依照字体外观来进行分类的方法。我们可以通过PANOSE体系将字体的外观特征进行整理，并且与其它字体归类比较。
		 * Panose的原形在1985年由Benjamin Bauermeister开发，当时一种字体由7位16进制数字定义，现在则发展为10位，也就是字体的十种特征。这每一位数字都给出了它定义的一种视觉外观的量度，如笔划的粗细或是字体衬线的样式等。
		 * Panose定义的范围：Latin Text，Latin Script，Latin Decorative，Iconographic，Japanese Text，Cyrillic Text，Hebrew。
		 * 	
		 * It is most likely to be suitable on Linux or OSX systems which don't have Microsoft's fonts installed.
		 * （它很可能适用于没有安装Microsoft字体的Linux或OSX系统。）
		 * 
		 * 1、获取Microsoft字体我们需要这些：a.在Microsoft平台上，嵌入PDF输出; b. docx4all - 所有平台 - 填充字体下拉列表
		 * setupMicrosoftFontFilenames();
		 * 2、 自动检测系统上可用的字体
		 * PhysicalFonts.discoverPhysicalFonts();
		 * 
		 */
		//Mapper fontMapper = new BestMatchingMapper();  
		/*
		 * 
		 * This mapper automatically maps document fonts for which the exact font is physically available.  
	     * Think of this as an identity mapping.  For  this reason, it will work best on Windows, or a system on 
	     * which Microsoft fonts have been installed.
	     * （此映射器自动映射确切可用的文档字体，将此视为标识映射；基于这个原因，它在Windows系统或安装了微软字体库的系统运行的更好。）
	     * You can manually add your own additional mappings if you wish.
	     * 如果需要，你可以手动添加自己的字体映射
		 * 
		 * 1、 自动检测系统上可用的字体
		 * PhysicalFonts.discoverPhysicalFonts();
		 * 
		 */
		Mapper fontMapper = new IdentityPlusMapper();
		//遍历自定义的字体库信息
		for (FontScheme fontScheme : COMPLIED_FONTSCHEME.values()) {
			//获取字体库
			PhysicalFont physicalFont = PhysicalFonts.get(fontScheme.getFontName());
			//分别设置字体名和别名对应的字体库
			fontMapper.put(fontScheme.getFontName(), physicalFont );
			fontMapper.put(fontScheme.getFontAlias(), physicalFont );
		}
		//进行中文字体兼容处理
		fontMapper.put("微软雅黑",PhysicalFonts.get("Microsoft Yahei"));
        fontMapper.put("黑体",PhysicalFonts.get("SimHei"));
        fontMapper.put("楷体",PhysicalFonts.get("KaiTi"));
        fontMapper.put("隶书", PhysicalFonts.get("LiSu"));
        fontMapper.put("宋体",PhysicalFonts.get("SimSun"));
        fontMapper.put("宋体扩展",PhysicalFonts.get("simsun-extB"));
        fontMapper.put("新宋体",PhysicalFonts.get("NSimSun"));
        fontMapper.put("仿宋",PhysicalFonts.get("FangSong"));
        fontMapper.put("仿宋_GB2312",PhysicalFonts.get("FangSong_GB2312"));
        fontMapper.put("幼圆",PhysicalFonts.get("YouYuan"));
        fontMapper.put("华文宋体",PhysicalFonts.get("STSong"));
        fontMapper.put("华文仿宋", PhysicalFonts.get("STFangsong"));
        fontMapper.put("华文中宋",PhysicalFonts.get("STZhongsong"));
        fontMapper.put("华文行楷", PhysicalFonts.get("STXingkai"));
        
        return fontMapper;
	}
	
	 
	
	public void destroy() {
		COMPLIED_FONTSCHEME.clear();
	}
	
}
