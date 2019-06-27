/**
 * Copyright (C), 2016-2019
 * FileName: FreemarkerService
 * Author:   yacnog_liu
 * Date:     2019/3/18 9:51
 * Description: Freemarker 模板解析器
 */
package com.freemarker.demo.doc.impl;


import com.alibaba.fastjson.JSONObject;
import com.freemarker.demo.doc.IDocument;
import com.freemarker.demo.vo.WordMsg;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import org.springframework.util.StringUtils;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.Map;

public class FreemarkerService implements IDocument {
    /*
     * 全局唯一示例
     */
    private static Configuration configuration = null;

    /*
     *  Default: Freemarker 模板包路径
     */
    private String TEMPLATE_PACKAGE = "/templates";

    /*
     *  Default: Freemarker 模板名称
     */
    private String TEMPLATE = "template.ftl";

    public FreemarkerService() {
    }

    /**
     * @param TEMPLATE_PACKAGE 模板路径 eg:"/com/kingdee/shr/template"
     * @param TEMPLATE         模板名称 eg:"template.ftl" or "template"
     */
    public FreemarkerService(String TEMPLATE_PACKAGE, String TEMPLATE) {
        this.TEMPLATE_PACKAGE = TEMPLATE_PACKAGE;
        this.TEMPLATE = TEMPLATE;
    }

    /**
     * @param TEMPLATE 模板名称 eg:"template.ftl" or "template"
     */
    public FreemarkerService(String TEMPLATE) {
        this.TEMPLATE = TEMPLATE;
    }

    /**
     * 导出word到本地
     *
     * @param dataMap  数据模型
     * @param filePath 文件路径
     * @return jsonStr
     */
    @Override
    public String exportDoc(Map<?, ?> dataMap, String filePath) {
        WordMsg msg = new WordMsg();
        msg.setStatus(false);

        Writer out;
        if (dataMap.isEmpty()) {
            msg.setErrCode(101);
            msg.setErrMsg("数据模型dataMap不能为空！");

            return JSONObject.toJSONString(msg);
        }

        if (StringUtils.isEmpty(filePath)) {
            msg.setErrCode(102);
            msg.setErrMsg("文件路径filePath不能为空！");

            return JSONObject.toJSONString(msg);
        }

        File outFile = new File(filePath);
        try {
            out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile), "utf-8"));
        } catch (UnsupportedEncodingException e) {
            msg.setErrCode(103);
            msg.setErrMsg("不支持该编码！");
            msg.setException(e);

            return JSONObject.toJSONString(msg);

        } catch (FileNotFoundException e1) {
            msg.setErrCode(104);
            msg.setErrMsg(filePath + "文件未找到！");
            msg.setException(e1);

            return JSONObject.toJSONString(msg);
        }

        // 加载Template模板
        Template template = getTemplate(TEMPLATE);
        try {
            if (template != null) {
                template.process(dataMap, out);

                out.close();

                msg.setStatus(true);
            }
        } catch (TemplateException e) {
            msg.setErrCode(105);
            msg.setErrMsg("执行模板发生错误，检查模板是否存在不存在的变量！");
            msg.setException(e);

            return JSONObject.toJSONString(msg);
        } catch (IOException e1) {
            msg.setErrCode(106);
            msg.setErrMsg("模板写入输入对象发生错误！");
            msg.setException(e1);

            return JSONObject.toJSONString(msg);
        }

        return JSONObject.toJSONString(msg);

    }

    /**
     * 导出word响应到服务端Response
     *
     * @param docName  文档名称
     * @param dataMap  数据模型
     * @param response HttpServletResponse
     * @return jsonStr
     */
    @Override
    public String exportDoc(String docName, Map<?, ?> dataMap, HttpServletResponse response) {
        WordMsg msg = new WordMsg();
        msg.setStatus(false);

        InputStream fis;
        ServletOutputStream outputStream;
        Writer out;

        if (StringUtils.isEmpty(docName)) {
            msg.setErrCode(107);
            msg.setErrMsg("文档名称docName不能为空！");

            return JSONObject.toJSONString(msg);
        }

        if (dataMap.isEmpty()) {
            msg.setErrCode(108);
            msg.setErrMsg("数据模型dataMap不能为空！");

            return JSONObject.toJSONString(msg);
        }

        if (response != null) {
            response.reset();
        }

        File outFile = new File(docName);
        try {
            out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile), "utf-8"));
        } catch (FileNotFoundException e) {
            msg.setErrCode(109);
            msg.setErrMsg("文件未找到！");
            msg.setException(e);

            return JSONObject.toJSONString(msg);
        } catch (UnsupportedEncodingException e1) {
            msg.setErrCode(110);
            msg.setErrMsg("不支持的编码：utf-8");
            msg.setException(e1);

            return JSONObject.toJSONString(msg);
        }

        // 加载Template模板
        Template template = getTemplate(TEMPLATE);
        try {
            if (template != null) {
                template.process(dataMap, out);
                out.close();
            }

        } catch (TemplateException e) {
            msg.setErrCode(111);
            msg.setErrMsg("执行模板发生错误，检查模板是否存在不存在的变量！");
            msg.setException(e);

            return JSONObject.toJSONString(msg);
        } catch (IOException e1) {
            msg.setErrCode(112);
            msg.setErrMsg("模板写入输出对象发生错误！");
            msg.setException(e1);

            return JSONObject.toJSONString(msg);
        }


        try {
            fis = new FileInputStream(outFile);
        } catch (FileNotFoundException e) {
            msg.setErrCode(113);
            msg.setErrMsg("文件未找到！");
            msg.setException(e);

            return JSONObject.toJSONString(msg);
        }

        response.setCharacterEncoding("utf-8");
        response.setContentType("application/msword");
        try {
            docName = new String(docName.getBytes("UTF-8"), "ISO-8859-1");
        } catch (UnsupportedEncodingException e) {
            msg.setErrCode(114);
            msg.setErrMsg("不支持的编码！");
            msg.setException(e);

            return JSONObject.toJSONString(msg);
        }

        response.setHeader("Content-disposition", "attachment;filename=" + docName + ".doc");
        try {
            outputStream = response.getOutputStream();
        } catch (IOException e) {
            msg.setErrCode(115);
            msg.setErrMsg("HttpServletResponse 获取输出流失败！");
            msg.setException(e);

            return JSONObject.toJSONString(msg);
        }

        byte[] buffer = new byte[512];  // 缓冲区
        int bytesToRead = -1;
        try {
            while ((bytesToRead = fis.read(buffer)) != -1) {
                outputStream.write(buffer, 0, bytesToRead);
            }

            msg.setStatus(true);
        } catch (IOException e) {
            msg.setErrCode(116);
            msg.setException(e);

            return JSONObject.toJSONString(msg);
        } finally {
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

            if (outFile != null) {
                //删除临时文件
                outFile.delete();
            }
        }

        return JSONObject.toJSONString(msg);

    }


    /**
     * 获取Freemarker Template
     *
     * @param templateName 模板名称  .ftl 结尾
     * @return Freemarker Template
     */
    private Template getTemplate(String templateName) {
        configuration = (configuration == null) ? getConfiguration() : configuration;

        if (StringUtils.isEmpty(templateName)) {
            return null;
        }

        // 模板名称处理 template template.ftl
        if (templateName.lastIndexOf(".ftl") == -1) {
            // 追加模板类型结尾
            templateName = templateName.concat(".ftl");
        }

        try {
            return configuration.getTemplate(templateName);

        } catch (IOException e) {
            System.out.println("***********************获取Freemaker模板发生错误：模板未找到或模板解析发生错误！" + e);
        }

        return null;


    }

    private Configuration getConfiguration() {
        Configuration configuration = FreemarkerHolder.configuration;
        configuration.setDefaultEncoding("utf-8");
        // 模板路径
        configuration.setClassForTemplateLoading(this.getClass(), TEMPLATE_PACKAGE);

        return configuration;
    }

    /**
     * 设计模式-单例模式
     */
    private static class FreemarkerHolder {

        private static Configuration configuration = new Configuration();

    }
}
