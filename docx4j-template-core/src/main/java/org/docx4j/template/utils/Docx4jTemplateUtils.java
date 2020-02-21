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
package org.docx4j.template.utils;


import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.Map;
import java.util.regex.Pattern;

import javax.servlet.http.HttpServletResponse;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Document;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
 
/**
 * 关于文件操作的工具类 （来自：https://blog.csdn.net/qq_35598240/article/details/84439929）
 * 备注：该工具只能解决固定模板的word生成
 * @author kaizen
 */
public final class Docx4jTemplateUtils {
 
    private static final Logger logger = LoggerFactory.getLogger(Docx4jUtils.class);
 
    /*
     * 替换变量并下载word文档
     *
     * @param inputStream
     * @param map
     * @param response
     * @param fileName
     */
    public static void downloadDocUseDoc4j(InputStream inputStream, Map<String, String> map,
                                          HttpServletResponse response, String fileName) {
 
        try {
            // 设置响应头
            fileName = URLEncoder.encode(fileName, "UTF-8");
            response.setContentType("application/octet-stream;charset=UTF-8");
            response.setCharacterEncoding("utf-8");
            response.setHeader("Content-Disposition", "attachment; filename=" + fileName + ".docx");
            response.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
 
            OutputStream outs  = response.getOutputStream();
            Docx4jTemplateUtils.replaceDocUseDoc4j(inputStream,map,outs);
 
        } catch (Exception e) {
            logger.error(e.getMessage(), e);
        }
    }
 
    /*
     * 替换变量并输出word文档
     * @param inputStream
     * @param map
     * @param outputStream
     */
    public static void replaceDocUseDoc4j(InputStream inputStream, Map<String, String> map,
                                          OutputStream outputStream) {
        try {
            WordprocessingMLPackage doc = WordprocessingMLPackage.load(inputStream);
            MainDocumentPart mainDocumentPart = doc.getMainDocumentPart();
            if (null != map && !map.isEmpty()) {
                // 将${}里的内容结构层次替换为一层
            	Docx4jTemplateUtils .cleanDocumentPart(mainDocumentPart);
                // 替换文本内容
                mainDocumentPart.variableReplace(map);
            }
 
            // 输出word文件
            doc.save(outputStream);
            outputStream.flush();
        } catch (Exception e) {
            logger.error(e.getMessage(), e);
        }
    }
 
 
    /*
     * cleanDocumentPart
     *
     * @param documentPart
     */
    public static boolean cleanDocumentPart(MainDocumentPart documentPart) throws Exception {
        if (documentPart == null) {
            return false;
        }
        Document document = documentPart.getContents();
        String wmlTemplate =
                XmlUtils.marshaltoString(document, true, false, Context.jc);
        document = (Document) XmlUtils.unwrap(DocxVariableClearUtils.doCleanDocumentPart(wmlTemplate, Context.jc));
        documentPart.setContents(document);
        return true;
    }
 
    /*
     * 清扫 docx4j 模板变量字符,通常以${variable}形式
     * <p>
     * XXX: 主要在上传模板时处理一下, 后续
     *
     * @author liliang
     * @since 2018-11-07
     */
    private static class DocxVariableClearUtils {
 
 
        /**
         * 去任意XML标签
         */
        private static final Pattern XML_PATTERN = Pattern.compile("<[^>]*>");
 
        private DocxVariableClearUtils() {
        }
 
        /**
         * start符号
         */
        private static final char PREFIX = '$';
 
        /**
         * 中包含
         */
        private static final char LEFT_BRACE = '{';
 
        /**
         * 结尾
         */
        private static final char RIGHT_BRACE = '}';
 
        /**
         * 未开始
         */
        private static final int NONE_START = -1;
 
        /**
         * 未开始
         */
        private static final int NONE_START_INDEX = -1;
 
        /**
         * 开始
         */
        private static final int PREFIX_STATUS = 1;
 
        /**
         * 左括号
         */
        private static final int LEFT_BRACE_STATUS = 2;
 
        /**
         * 右括号
         */
        private static final int RIGHT_BRACE_STATUS = 3;
 
 
        /*
         * doCleanDocumentPart
         *
         * @param wmlTemplate
         * @param jc
         * @return
         * @throws JAXBException
         */
        private static Object doCleanDocumentPart(String wmlTemplate, JAXBContext jc) throws JAXBException {
            // 进入变量块位置
            int curStatus = NONE_START;
            // 开始位置
            int keyStartIndex = NONE_START_INDEX;
            // 当前位置
            int curIndex = 0;
            char[] textCharacters = wmlTemplate.toCharArray();
            StringBuilder documentBuilder = new StringBuilder(textCharacters.length);
            documentBuilder.append(textCharacters);
            // 新文档
            StringBuilder newDocumentBuilder = new StringBuilder(textCharacters.length);
            // 最后一次写位置
            int lastWriteIndex = 0;
            for (char c : textCharacters) {
                switch (c) {
                    case PREFIX:
                        // TODO 不管其何状态直接修改指针,这也意味着变量名称里面不能有PREFIX
                        keyStartIndex = curIndex;
                        curStatus = PREFIX_STATUS;
                        break;
                    case LEFT_BRACE:
                        if (curStatus == PREFIX_STATUS) {
                            curStatus = LEFT_BRACE_STATUS;
                        }
                        break;
                    case RIGHT_BRACE:
                        if (curStatus == LEFT_BRACE_STATUS) {
                            // 接上之前的字符
                            newDocumentBuilder.append(documentBuilder.substring(lastWriteIndex, keyStartIndex));
                            // 结束位置
                            int keyEndIndex = curIndex + 1;
                            // 替换
                            String rawKey = documentBuilder.substring(keyStartIndex, keyEndIndex);
                            // 干掉多余标签
                            String mappingKey = XML_PATTERN.matcher(rawKey).replaceAll("");
                            if (!mappingKey.equals(rawKey)) {
                                char[] rawKeyChars = rawKey.toCharArray();
                                // 保留原格式
                                StringBuilder rawStringBuilder = new StringBuilder(rawKey.length());
                                // 去掉变量引用字符
                                for (char rawChar : rawKeyChars) {
                                    if (rawChar == PREFIX || rawChar == LEFT_BRACE || rawChar == RIGHT_BRACE) {
                                        continue;
                                    }
                                    rawStringBuilder.append(rawChar);
                                }
                                // FIXME 要求变量连在一起
                                String variable = mappingKey.substring(2, mappingKey.length() - 1);
                                int variableStart = rawStringBuilder.indexOf(variable);
                                if (variableStart > 0) {
                                    rawStringBuilder = rawStringBuilder.replace(variableStart, variableStart + variable.length(), mappingKey);
                                }
                                newDocumentBuilder.append(rawStringBuilder.toString());
                            } else {
                                newDocumentBuilder.append(mappingKey);
                            }
                            lastWriteIndex = keyEndIndex;
 
                            curStatus = NONE_START;
                            keyStartIndex = NONE_START_INDEX;
                        }
                    default:
                        break;
                }
                curIndex++;
            }
            // 余部
            if (lastWriteIndex < documentBuilder.length()) {
                newDocumentBuilder.append(documentBuilder.substring(lastWriteIndex));
            }
            return XmlUtils.unmarshalString(newDocumentBuilder.toString(), jc);
        }
 
    }
}