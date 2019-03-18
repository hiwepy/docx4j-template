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

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.template.io.WordprocessingMLPackageRender;
import org.docx4j.template.utils.WMLPackageUtils;

public class AddingAnInlineImageTest {

	
	/** 
     *  像往常一样, 我们创建了一个包(package)来容纳文档. 
     *  然后我们创建了一个指向将要添加到文档的图片的文件对象.为了能够对图片做一些操作, 我们将它转换 
     *  为字节数组. 最后我们将图片添加到包中并保存这个包(package). 
     */  
    public static void main (String[] args) throws Exception {  
        WordprocessingMLPackage  wordMLPackage = WordprocessingMLPackage.createPackage();  
   
        File file = new File("src/main/resources/iProfsLogo.png");  
        byte[] bytes = WMLPackageUtils.imageToByteArray(file);
        
        WordprocessingMLPackageRender render = new WordprocessingMLPackageRender(wordMLPackage);
        
        render.addImageToPackage(bytes);
   
        wordMLPackage.save(new java.io.File("src/main/files/HelloWord7.docx"));  
    }  
   
    
}
