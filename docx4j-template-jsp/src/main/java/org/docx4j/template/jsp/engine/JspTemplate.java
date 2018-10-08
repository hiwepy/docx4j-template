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
package org.docx4j.template.jsp.engine;

import java.io.IOException;
import java.io.OutputStream;
import java.io.Writer;
import java.util.Map;

import javax.servlet.ServletException;

/**
 * 
 * @className	： JspTemplate
 * @description	： TODO(描述这个类的作用)
 * @author 		： <a href="https://github.com/vindell">vindell</a>
 * @date		： 2017年5月24日 下午10:26:02
 * @version 	V1.0
 */
public interface JspTemplate {

	public String getName();
	
	public void render(String requestURL,Map<String, Object> variables, Writer out) throws IOException, ServletException;
	
	public void render(String requestURL,Map<String, Object> variables, OutputStream out) throws IOException, ServletException;
	
}
