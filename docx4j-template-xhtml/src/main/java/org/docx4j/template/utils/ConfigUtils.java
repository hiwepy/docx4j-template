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

import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

/**
 * 
 * @className	： ConfigUtils
 * @description	： TODO(描述这个类的作用)
 * @author 		： <a href="https://github.com/vindell">vindell</a>
 * @date		： 2017年5月24日 下午10:27:22
 * @version 	V1.0
 */
public class ConfigUtils {

	private static final String PLUS = "+";

	private static final String MINUS = "-";
	
	public static Properties filterWithPrefix(String filter,String prefix,Properties input, boolean escape) {
		Properties ret = new Properties();
		for (Object entry : input.keySet()) {
			String key = entry.toString();
			if (escape) {
				key = key.replace('_', '.');
			}
			if(key.startsWith(filter)) {
				String value = input.getProperty(key);
				if (escape) {
					if (value.startsWith(PLUS)) {
						key += PLUS;
						value = value.substring(PLUS.length());
					} else if (value.startsWith(MINUS)) {
						key += MINUS;
						value = value.substring(MINUS.length());
					}
				}
				key = key.substring(prefix.length());
				ret.setProperty(key, value);
			}
		}
		return ret;
	}
	
	public static Map<String, String> filterWithPrefix(String prefix, Map<String, String> input, boolean escape) {
		Map<String, String> ret = new HashMap<String, String>();
		for (Map.Entry<String, String> entry : input.entrySet()) {
			String key = entry.getKey();
			if (escape) {
				key = key.replace('_', '.');
			}
			if(key.startsWith(prefix)) {
				key = key.substring(prefix.length());
				String value = entry.getValue();
				if (escape) {
					if (value.startsWith(PLUS)) {
						key += PLUS;
						value = value.substring(PLUS.length());
					} else if (value.startsWith(MINUS)) {
						key += MINUS;
						value = value.substring(MINUS.length());
					}
				}
				ret.put(key, value);
			}
		}
		return ret;
	}
	
}
