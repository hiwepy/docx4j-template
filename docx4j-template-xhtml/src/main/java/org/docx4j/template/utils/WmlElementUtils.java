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

import java.util.List;

import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.Ftr;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcPr;
import org.docx4j.wml.TcPrInner.GridSpan;
import org.docx4j.wml.Text;

/**
 * @className	： WmlElementUtils
 * @description	： TODO(描述这个类的作用)
 * @author 		： <a href="https://github.com/vindell">vindell</a>
 * @date		： 2017年5月24日 下午10:27:58
 * @version 	V1.0
 */
public final class WmlElementUtils {

	protected static ObjectFactory factory = Context.getWmlObjectFactory();
	
	/**
     *  First we create a footer, a paragraph, a run and a text. We add the given
     *  given content to the text and add that to the run. The run is then added to
     *  the paragraph, which is in turn added to the footer. Finally we return the
     *  footer.
     *  @param content
     *  @return
     */
	public static Ftr createFooter(String content) {
        Ftr footer = factory.createFtr();
        P paragraph = factory.createP();
        R run = factory.createR();
        Text text = new Text();
        text.setValue(content);
        run.getContent().add(text);
        paragraph.getContent().add(run);
        footer.getContent().add(paragraph);
        return footer;
    }

    /** 
     *  创建一个对象工厂并用它创建一个段落和一个可运行块R. 
     *  然后将可运行块添加到段落中. 接下来创建一个图画并将其添加到可运行块R中. 最后我们将内联 
     *  对象添加到图画中并返回段落对象. 
     * 
     * @param   inline 包含图片的内联对象. 
     * @return  包含图片的段落 
     */  
    public static P addInlineImageToParagraph(Inline inline) {
    	ObjectFactory factory = new ObjectFactory();
        // 添加内联对象到一个段落中  
        P paragraph = factory.createP();  
        R run = factory.createR();  
        paragraph.getContent().add(run);  
        Drawing drawing = factory.createDrawing();  
        run.getContent().add(drawing);  
        drawing.getAnchorOrInline().add(inline);  
        return paragraph;  
    } 
	
    /*
     * http://53873039oycg.iteye.com/blog/2194849
     *  实现思路:
        主要分在当前行上方插入行和在当前行下方插入行。对首尾2行特殊处理，在有跨行合并情况时，在第一行上面或者在最后一行下面插入是不会跨行的但是可能会跨列。
       对于中间的行，主要参照当前行，如果当前行跨行，则新增行也跨行，如果当前行单元格结束跨行，则新增的上方插入行跨行，下方插入行不跨行，如果当前行单元格开始跨行，则新增的上方插入行不跨行，下发插入行跨行。
       主要思路就是这样，插入的时候需要得到真实位置的单元格，代码如下:
       
     */
    // 按位置得到单元格(考虑跨列合并情况)  
    public Tc getTcByPosition(List<Tc> tcList, int position) {  
        int k = 0;  
        for (int i = 0, len = tcList.size(); i < len; i++) {  
            Tc tc = tcList.get(i);  
            TcPr trPr = tc.getTcPr();  
            if (trPr != null) {  
                GridSpan gridSpan = trPr.getGridSpan();  
                if (gridSpan != null) {  
                    k += gridSpan.getVal().intValue() - 1;  
                }
            }  
            if (k >= position) {  
                return tcList.get(i);  
            }  
            k++;  
        }  
        if (position < tcList.size()) {  
            return tcList.get(position);  
        }  
        return null;  
    }  
}
