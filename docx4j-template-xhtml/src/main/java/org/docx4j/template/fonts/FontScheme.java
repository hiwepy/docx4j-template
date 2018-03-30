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

import java.io.Serializable;
import java.net.URL;

import org.docx4j.template.utils.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 
 * @className	： FontScheme
 * @description	： TODO(描述这个类的作用)
 * @author 		： <a href="https://github.com/vindell">vindell</a>
 * @date		： 2017年5月24日 下午10:26:18
 * @version 	V1.0
 */
@SuppressWarnings("serial")
public class FontScheme implements Serializable {

	protected static Logger LOG = LoggerFactory.getLogger(FontScheme.class);
	protected String fontName;
	protected String fontAlias;
	protected URL fontURL;
	
	public FontScheme(String fontName, String fontAlias, URL fontURL) {
		this.fontName = fontName;
		this.fontAlias = fontAlias;
		this.fontURL = fontURL;
	}

	public String getFontName() {
		return fontName;
	}

	public String getFontAlias() {
		return fontAlias;
	}

	public URL getFontURL() {
		return fontURL;
	}

	@Override
	public String toString() {
		return "PhysicalFont [" +  StringUtils.join(new Object[]{ fontName , fontAlias , fontURL }, "-") + " ]";
	}

}
