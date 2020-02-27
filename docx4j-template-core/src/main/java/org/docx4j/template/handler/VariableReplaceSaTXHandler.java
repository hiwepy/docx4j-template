/** 
 * Copyright (C) 2018 Jeebiz (http://jeebiz.net).
 * All Rights Reserved. 
 */
package org.docx4j.template.handler;

import java.util.Map;

import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import javax.xml.stream.XMLStreamWriter;

import org.docx4j.openpackaging.parts.StAXHandlerAbstract;
import org.xml.sax.SAXException;

public class VariableReplaceSaTXHandler extends StAXHandlerAbstract {
	/**
	 * 变量占位符开始位，默认：${
	 */
	protected String placeholderStart = "${";
	/**
	 * 变量占位符结束位，默认：}
	 */
	protected String placeholderEnd = "}";
	protected Map<String, Object> variables;

	public VariableReplaceSaTXHandler(Map<String, Object> variables) throws SAXException {
		super();
		this.variables = variables;
	}

	public VariableReplaceSaTXHandler(String placeholderStart, String placeholderEnd, Map<String, Object> variables)
			throws SAXException {
		super();
		this.placeholderStart = placeholderStart;
		this.placeholderEnd = placeholderEnd;
		this.variables = variables;
	}

	@Override
	public void handleCharacters(XMLStreamReader xmlr, XMLStreamWriter writer) throws XMLStreamException {

		StringBuilder sb = new StringBuilder();
		sb.append(xmlr.getTextCharacters(), xmlr.getTextStart(), xmlr.getTextLength());

		String wmlString = replace(sb.toString(), 0, new StringBuilder(), variables).toString();
//		System.out.println(wmlString);

		char[] charOut = wmlString.toCharArray();
		writer.writeCharacters(charOut, 0, charOut.length);

//		writer.writeCharacters(xmlr.getTextCharacters(),
//				xmlr.getTextStart(), xmlr.getTextLength());

	}

	private StringBuilder replace(String wmlTemplateString, int offset, StringBuilder strB,
			Map<String, Object> mappings) {

		int startKey = wmlTemplateString.indexOf(placeholderStart, offset);
		if (startKey == -1)
			return strB.append(wmlTemplateString.substring(offset));
		else {
			strB.append(wmlTemplateString.substring(offset, startKey));
			int keyEnd = wmlTemplateString.indexOf(placeholderEnd, startKey);
			if (keyEnd > 0) {
				String key = wmlTemplateString.substring(startKey + 2, keyEnd);
				Object val = mappings.get(key);
				if (val == null) {
					System.out.println("Invalid key '" + key + "' or key not mapped to a value");
					strB.append(key);
				} else {
					strB.append(val.toString());
				}
				return replace(wmlTemplateString, keyEnd + 1, strB, mappings);
			} else {
				System.out.println("Invalid key: could not find '}' ");
				strB.append("$");
				return replace(wmlTemplateString, offset + 1, strB, mappings);
			}
		}
	}
}