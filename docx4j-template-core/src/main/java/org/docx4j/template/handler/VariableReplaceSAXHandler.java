/** 
 * Copyright (C) 2018 Jeebiz (http://jeebiz.net).
 * All Rights Reserved. 
 */
package org.docx4j.template.handler;

import java.util.Map;

import org.docx4j.openpackaging.parts.SAXHandler;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;

public class VariableReplaceSAXHandler extends SAXHandler implements ContentHandler {
	/**
	 * 变量占位符开始位，默认：${
	 */
	protected String placeholderStart = "${";
	/**
	 * 变量占位符结束位，默认：}
	 */
	protected String placeholderEnd = "}";
	protected Map<String, Object> variables;

	public VariableReplaceSAXHandler(Map<String, Object> variables) throws SAXException {
		super();
	}
	
	public VariableReplaceSAXHandler(String placeholderStart, String placeholderEnd ,Map<String, Object> variables) throws SAXException {
		super();
		this.placeholderStart = placeholderStart;
		this.placeholderEnd = placeholderEnd;
		this.variables = variables;
	}

	@Override
	public void characters(char[] ch, int start, int length) throws SAXException {

		StringBuilder sb = new StringBuilder();
		sb.append(ch, start, length);

		String wmlString = replace(sb.toString(), 0, new StringBuilder(), variables).toString();
//		System.out.println(wmlString);

		char[] charOut = wmlString.toCharArray();
		
		this.getContentHandler().characters(charOut, 0, charOut.length);

	}

	private StringBuilder replace(String wmlTemplateString, int offset, StringBuilder strB,
			Map<String, Object> mappings) {

		int startKey = wmlTemplateString.indexOf(placeholderStart, offset);
		if (startKey == -1)
			return strB.append(wmlTemplateString.substring(offset));
		else {
			strB.append(wmlTemplateString.substring(offset, startKey));
			int keyEnd = wmlTemplateString.indexOf(placeholderEnd, startKey);
			String key = wmlTemplateString.substring(startKey + 2, keyEnd);
			Object val = mappings.get(key);
			if (val == null) {
				System.out.println("Invalid key '" + key + "' or key not mapped to a value");
				strB.append(key);
			} else {
				strB.append(val.toString());
			}
			return replace(wmlTemplateString, keyEnd + 1, strB, mappings);
		}
	}
}