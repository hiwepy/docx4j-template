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

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.xml.bind.JAXBException;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.AlternativeFormatInputPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.CTAltChunk;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;

/**
 * TODO
 * @author <a href="https://github.com/hiwepy">hiwepy</a>
 */
@SuppressWarnings("unchecked")
public class WMLPackageUtils {
	
	protected static String CONTENT_TYPE = "";
	
	/*
	 * 首先，我们创建一个可用作模版的简单的word文档。对于此只需打开Word，创建新文档然后保存为template.docx，这就是我们将要用于添加内容的word文档。
	 * 我们需要做的第一件事是使用docx4j将这个文档加载进来，你可以使用下面的几行代码做这件事：
	 */
	public static WordprocessingMLPackage getWMLPackageTemplate(String filepath) throws Docx4JException, FileNotFoundException {
		WordprocessingMLPackage template = WordprocessingMLPackage.load(new FileInputStream(new File(filepath)));
		return template;
	}
	
	public static <T> List<T> getChildrenElements(Object source, Class<T> targetClass) {  
        List<T> result = new ArrayList<T>();  
        //获取真实的对象
    	Object target = XmlUtils.unwrap(source);
    	//if (target.getClass().equals(targetClass)) {
        if (targetClass.isAssignableFrom(target.getClass())) { 
            result.add((T)target);
        } else if (target instanceof ContentAccessor) {  
            List<?> children = ((ContentAccessor) target).getContent();  
            //if (children.getClass().equals(targetClass)) {
            if (targetClass.isAssignableFrom(children.getClass())) { 
                result.add((T)children);
            }
        }
        return result;  
    }
	
	/*
	 * 这样会返回一个表示完整的空白（在此时）文档Java对象。现在我们可以使用Docx4J API添加、删除以及更新这个word文档的内容，Docx4J有一些你可以用于遍历该文档的工具类。
	 * 我自己写了几个助手方法使查找指定占位符并用真实内容进行替换的操作变地很简单。让我们来看一下其中的一个，这个计算是几个JAXB计算的包装器，
	 * 允许你针对一个特定的类来搜索指定元素以及它所有的孩子，例如，你可以用它获取文档中所有的表格、表格中所有的行以及其它类似的操作。 
	 */
	public static <T> List<T> getTargetElements(Object source, Class<T> targetClass) {  
        List<T> result = new ArrayList<T>();  
        //获取真实的对象
    	Object target = XmlUtils.unwrap(source);
        //if (target.getClass().equals(targetClass)) {
    	if (targetClass.isAssignableFrom(target.getClass())) { 
            result.add((T) target);  
        } else if (target instanceof ContentAccessor) {  
            List<?> children = ((ContentAccessor) target).getContent();  
            for (Object child : children) {  
                result.addAll(getTargetElements(child, targetClass));  
            }
        }
        return result;  
    }  
	
	/*
	 * 例如你动态设置一个文档的标题。首先，在前面创建的模版文档中添加一个自定义占位符，我使用SJ_EX1作为占位符，我们将要用name参数来替换这个值。
	 * 在docx4j中基本的文本元素用org.docx4j.wml.Text类来表示，替换这个简单的占位符我们需要做的就是调用这个方法： 
	 */
	public static void replacePlaceholder(WordprocessingMLPackage template, String name, String placeholder ) { 
		//获取文档对象中所有的Text类型对象
	    List<Text> texts = getTargetElements(template.getMainDocumentPart(), Text.class);  
	   //循环Text类型对象集合
	    for (Text text : texts) {  
	        Text textElement = (Text) text;  
	        if (textElement.getValue().equals(placeholder)) {  
	            textElement.setValue(name);  
	        }  
	    }  
	} 
	
	public static void writeDocxToStream(WordprocessingMLPackage template, String target) throws IOException, Docx4JException {  
        File f = new File(target);  
        template.save(f);  
    }  
	
	/*
	 *  向模版文档添加段落
		你可能想知道为什么我们需要添加段落？我们已经可以添加文本，难道段落不就是一大段的文本吗？好吧，既是也不是，一个段落确实看起来像是一大段文本，但你需要考虑的是换行符，如果你像前面一样添加一个Text元素并且在文本中添加换行符，它们并不会出现，当你想要换行符时，你就需要创建一个新的段落。然而，幸运的是这对于Docx4j来说也非常地容易。
		做这个需要下面的几步：
		
		    从模版中找到要替换的段落
		    将输入文本拆分成单独的行
		    每一行基于模版中的段落创建一个新的段落
		    移除原来的段落
    
	 * ******************************************************************
	 */
	public static void replaceParagraph(String placeholder, String textToAdd, WordprocessingMLPackage template, ContentAccessor addTo) {  
        // 1. get the paragraph  
        List<P> paragraphs = getTargetElements(template.getMainDocumentPart(), P.class);  
        P toReplace = null;  
        for (P p : paragraphs) {  
            List<Text> texts = getTargetElements(p, Text.class);  
            for (Text t : texts) {  
                Text content = (Text) t;  
                if (content.getValue().equals(placeholder)) {  
                    toReplace = (P) p;  
                    break;  
                }  
            }  
        }  
      
        // we now have the paragraph that contains our placeholder: toReplace  
        // 2. split into seperate lines  
        String as[] = StringUtils.splitPreserveAllTokens(textToAdd, '\n');  
      
        for (int i = 0; i < as.length; i++) {  
            String ptext = as[i];  
      
            // 3. copy the found paragraph to keep styling correct  
            P copy = (P) XmlUtils.deepCopy(toReplace);  
      
            // replace the text elements from the copy  
            List<?> texts = getTargetElements(copy, Text.class);  
            if (texts.size() > 0) {  
                Text textToReplace = (Text) texts.get(0);  
                textToReplace.setValue(ptext);  
            }  
      
            // add the paragraph to the document  
            addTo.getContent().add(copy);  
        }  
      
        // 4. remove the original one  
        ((ContentAccessor)toReplace.getParent()).getContent().remove(toReplace);  
      
    }  
	
	/*
	 * 该方法找到表格，获取第一行并且遍历提供的map向表格添加新行，在将其返回之前删除模版行。这个方法用到了两个助手方法：addRowToTable 和 getTemplateTable。我们首先看一下后面的那个： 
	 */
	public static Tbl getTable(List<Tbl> tables, String placeholder) throws Docx4JException {  
        for (Iterator<Tbl> iterator = tables.iterator(); iterator.hasNext();) {  
        	Tbl tbl = iterator.next();
        	//查找当前table下面的text对象
            List<Text> textElements = getTargetElements(tbl, Text.class);  
            for (Text text : textElements) {
                Text textElement = (Text) text;
                //
                if (textElement.getValue() != null && textElement.getValue().equals(placeholder))  {
                    return (Tbl) tbl;  
                }
            }  
        }  
        return null;  
    }  
    
	public static void replaceTable(String[] placeholders, List<Map<String, String>> textToAdd,  
            WordprocessingMLPackage template) throws Docx4JException, JAXBException {  
	    List<Tbl> tables = getTargetElements(template.getMainDocumentPart(), Tbl.class);  
	  
	    // 1. find the table  
	    Tbl tempTable = getTable(tables, placeholders[0]);  
	    List<Tr> rows = getTargetElements(tempTable, Tr.class);  
	  
	    // first row is header, second row is content  
	    if (rows.size() == 2) {  
	    	
	        // this is our template row  
	        Tr templateRow = (Tr) rows.get(1);  
	  
	        for (Map<String, String> replacements : textToAdd) {  
	            // 2 and 3 are done in this method  
	            addRowToTable(tempTable, templateRow, replacements);  
	        }  
	  
	        // 4. remove the template row  
	        tempTable.getContent().remove(templateRow);  
	    }  
	}  
    
	public static void addRowToTable(Tbl reviewtable, Tr templateRow, Map<String, String> replacements) {  
        Tr workingRow = (Tr) XmlUtils.deepCopy(templateRow);  
        List<?> textElements = getTargetElements(workingRow, Text.class);  
        for (Object object : textElements) {  
            Text text = (Text) object;  
            String replacementValue = (String) replacements.get(text.getValue());  
            if (replacementValue != null)  
                text.setValue(replacementValue);  
        }  
      
        reviewtable.getContent().add(workingRow);  
    } 
	
    public static InputStream mergeDocx(final List<InputStream> streams)  throws Docx4JException, IOException {  
      
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
                    insertDocx(target.getMainDocumentPart(),IOUtils.toByteArray(is), chunkId++);  
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
    private static void insertDocx(MainDocumentPart main, byte[] bytes, int chunkId) {  
        try {  
            AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(new PartName("/part" + chunkId + ".docx"));  
            afiPart.setContentType(new ContentType(CONTENT_TYPE));  
            afiPart.setBinaryData(bytes);  
            Relationship altChunkRel = main.addTargetPart(afiPart);  
      
            CTAltChunk chunk = Context.getWmlObjectFactory().createCTAltChunk();  
            chunk.setId(altChunkRel.getId());  
      
            main.addObject(chunk);  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
    } 
    
	/*
     * 将图片从文件对象转换成字节数组. 
     * @param file  将要转换的文件 
     * @return      包含图片字节数据的字节数组 
     * @throws FileNotFoundException 
     * @throws IOException 
     */  
	public static byte[] imageToByteArray(File file) throws FileNotFoundException, IOException {  
        InputStream is = new FileInputStream(file );  
        long length = file.length();  
        // 不能使用long类型创建数组, 需要用int类型.  
        if (length > Integer.MAX_VALUE) {  
            System.out.println("File too large!!");  
        }  
        byte[] bytes = new byte[(int)length];  
        int offset = 0;  
        int numRead = 0;  
        while (offset < bytes.length && (numRead=is.read(bytes, offset, bytes.length-offset)) >= 0) {  
            offset += numRead;  
        }  
        // 确认所有的字节都没读取  
        if (offset < bytes.length) {  
            System.out.println("Could not completely read file " +file.getName());  
        }  
        is.close();  
        return bytes;  
    }   
	
}
