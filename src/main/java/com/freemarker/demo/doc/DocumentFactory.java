/**
 * Copyright (C), 2016-2019
 * FileName: DocumentFactory
 * Author:   yacnog_liu
 * Date:     2019/3/19 8:43
 * Description: DocumentFactory
 * <p>
 * 设计模式-静态工厂方法模式
 */
package com.freemarker.demo.doc;


import com.freemarker.demo.doc.impl.FreemarkerService;

public class DocumentFactory {

    /**
     * 获取FreemarkerService实例 template默认模板
     *
     * @return IDocument
     */
    public static IDocument produceFreemarker() {
        return new FreemarkerService();
    }

    /**
     * @param templatePackage 模板路径 eg:"/com/kingdee/shr/template"
     * @param templateName    模板名称 eg:"template.ftl" or "template"
     * @return IDocument
     */
    public static IDocument produceFreemarker(String templatePackage, String templateName) {
        return new FreemarkerService(templatePackage, templateName);
    }

    /**
     * @param templateName 模板名称 eg:"template.ftl" or "template"
     * @return
     */
    public static IDocument produceFreemarker(String templateName) {
        return new FreemarkerService(templateName);
    }
}
