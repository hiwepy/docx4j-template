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

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.io.IOUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.contenttype.ContentTypes;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.AlternativeFormatInputPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.template.io.WordprocessingMLPackageBuilder;
import org.docx4j.template.io.WordprocessingMLPackageWriter;
import org.docx4j.wml.CTAltChunk;

/**
 * 
 * @className	： Docx4jUtils
 * @description	： TODO(描述这个类的作用)
 * @author 		： <a href="https://github.com/vindell">vindell</a>
 * @date		： 2017年5月24日 下午10:27:30
 * @version 	V1.0
 */
public class Docx4jUtils {

	protected static WordprocessingMLPackageBuilder WMLPACKAGE_BUILDER = new WordprocessingMLPackageBuilder();
	protected static WordprocessingMLPackageWriter WMLPACKAGE_WRITER = new WordprocessingMLPackageWriter();

	/**
	 * 生成临时文件位置
	 */
	public static String getTempPath() {
		return System.getProperty("java.io.tmpdir") + File.separator + System.currentTimeMillis();
	}

	/**
	 * docx文档转换为PDF
	 * 
	 * @param docx
	 *            docx文档
	 * @param pdfPath
	 *            PDF文档存储路径
	 * @throws Exception
	 *             可能为Docx4JException, FileNotFoundException, IOException等
	 */
	public static void docxToPdf(String docxPath, String pdfPath) throws Exception {
		OutputStream output = null;
		try {
			output = new FileOutputStream(pdfPath);
			
			WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.load(new File(docxPath));

			WMLPACKAGE_BUILDER.configChineseFonts(wmlPackage).configSimSunFont(wmlPackage);
			WMLPACKAGE_WRITER.writeToPDF(wmlPackage, output);

		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			IOUtils.closeQuietly(output);
		}
	}

	/**
	 * 把docx转成html
	 * @param docxFilePath
	 * @param htmlPath
	 * @throws Exception
	 */
	public static void docxToHtml(String docxFilePath, String htmlPath) throws Exception {
		OutputStream output = null;
		try {
			//
			WordprocessingMLPackage wmlPackage = WordprocessingMLPackage.load(new File(docxFilePath));

			WMLPACKAGE_BUILDER.configChineseFonts(wmlPackage).configSimSunFont(wmlPackage);
			
			WMLPACKAGE_WRITER.writeToHtml(wmlPackage, htmlPath);

		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			IOUtils.closeQuietly(output);
		}

	}
	

    public InputStream mergeDocx(final List<InputStream> streams)  throws Docx4JException, IOException {  
      
        WordprocessingMLPackage target = null;  
        final File generated = File.createTempFile("generated", ".docx");  
      
        int chunkId = 0;  
        Iterator<InputStream> it = streams.iterator();  
        while (it.hasNext()) {  
            InputStream is = it.next();  
            if (is != null) {  
                if (target == null) {  
                    // Copy first (master) document   
                    OutputStream os = new FileOutputStream(generated);  
                    os.write(IOUtils.toByteArray(is));  
                    os.close();  
      
                    target = WordprocessingMLPackage.load(generated);  
                } else {  
                    // Attach the others (Alternative input parts)   
                    insertDocx(target.getMainDocumentPart(),  
                            IOUtils.toByteArray(is), chunkId++);  
                }  
            }  
        }  
      
        if (target != null) {  
            target.save(generated);  
            return new FileInputStream(generated);  
        } else {  
            return null;  
        }  
    }  
      
    // 插入文档   
    private void insertDocx(MainDocumentPart main, byte[] bytes, int chunkId) {  
        try {  
            AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(new PartName("/part" + chunkId + ".docx"));  
            // afiPart.setContentType(new ContentType(CONTENT_TYPE));
            afiPart.setContentType(new ContentType(ContentTypes.APPLICATION_XML));
            afiPart.setBinaryData(bytes);  
            Relationship altChunkRel = main.addTargetPart(afiPart);  
      
            CTAltChunk chunk = Context.getWmlObjectFactory().createCTAltChunk();  
            chunk.setId(altChunkRel.getId());  
      
            main.addObject(chunk);  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
    } 
    

}
