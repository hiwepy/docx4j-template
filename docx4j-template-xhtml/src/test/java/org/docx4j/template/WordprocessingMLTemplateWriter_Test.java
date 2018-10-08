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
package org.docx4j.template;

import java.io.File;

import org.docx4j.template.io.WordprocessingMLTemplateWriter;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

public class WordprocessingMLTemplateWriter_Test {

	protected WordprocessingMLTemplateWriter wemplateWriter = null;
	
	@Before
	public void Before() {
		wemplateWriter = new WordprocessingMLTemplateWriter();
		
		//org.eclipse.persistence.jaxb.JAXBContextFactory
	}
	
	@Test
	public void test() throws Exception {
		
		System.out.println(wemplateWriter.writeToString(new File("D:\\HelloWord14.docx")));
		
	}
	
	@After
	public void after() {
		wemplateWriter = null;
	}
	
	
}
