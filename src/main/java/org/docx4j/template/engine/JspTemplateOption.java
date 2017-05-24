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
package org.docx4j.template.engine;

import java.util.Map;

import jetbrick.template.parser.ParserContext;

/**
 * 
 * @className	： JspTemplateOption
 * @description	： TODO(描述这个类的作用)
 * @author 		： <a href="https://github.com/vindell">vindell</a>
 * @date		： 2017年5月24日 下午10:26:12
 * @version 	V1.0
 */
public final class JspTemplateOption {
    private final boolean strict;
    private final boolean safecall;
    private final boolean trimLeadingWhitespaces;
    private final Map<String, Class<?>> symbols;

    protected JspTemplateOption(ParserContext ctx) {
        this.strict = ctx.isStrict();
        this.safecall = ctx.isSafecall();
        this.trimLeadingWhitespaces = ctx.isTrimLeadingWhitespaces();
        this.symbols = ctx.getSymbols();
    }

    public boolean isStrict() {
        return strict;
    }

    public boolean isSafecall() {
        return safecall;
    }

    public boolean isTrimLeadingWhitespaces() {
        return trimLeadingWhitespaces;
    }

    public Map<String, Class<?>> getSymbols() {
        return symbols;
    }
}
