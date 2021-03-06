/*
 * Copyright (c) 2018, hiwepy (https://github.com/hiwepy).
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

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.template.WordprocessingMLDocxStAXTemplate;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

public class WordprocessingMLDocxStAXTemplate_Test {

	protected WordprocessingMLDocxStAXTemplate docxTemplate = null;
	
	@Before
	public void Before() throws IOException {
		docxTemplate = new WordprocessingMLDocxStAXTemplate();
	}
	
	@Test
	public void test() throws Exception {
		Map<String, Object> variables = new HashMap<String, Object>();
		
		variables.put("title", "变量替换测试");
		variables.put("content", "测试效果不错");
		
		Map<String, String> map = new HashMap<>();
		map.put("title", "Ognl Expression 测试！ ");
		variables.put("map", map);
		
		File sourceDocx = new java.io.File("src/test/resources/tpl/template.docx");
		File outputDocx = new java.io.File("src/test/resources/output/docxTemplate_output3.docx");

		WordprocessingMLPackage wmlPackage = docxTemplate.process(sourceDocx, variables);
		wmlPackage.save(outputDocx);
		
		
	}
	
	@After
	public void after() {
		docxTemplate = null;
	}
	
}
