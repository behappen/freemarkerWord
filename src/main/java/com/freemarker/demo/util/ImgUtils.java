/**
 * Copyright (C), 2016-2019
 * FileName: ImgUtils
 * Author:   yacnog_liu
 * Date:     2019/3/21 8:55
 * Description: Img工具类
 */
package com.freemarker.demo.util;

import org.springframework.util.StringUtils;
import sun.misc.BASE64Encoder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ImgUtils {

    /**
     * 通过图片路径获取其 Base64 码
     *
     * @param imgPath @Not-Null
     * @return img_Base64code
     */
    public static String getImgBase64Code(String imgPath) throws IOException {

        if (StringUtils.isEmpty(imgPath)) {
            return null;
        }

        File file = new File(imgPath);

        FileInputStream fis = null;
        byte[] data = null;
        try {
            fis = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            System.out.println("*****************ImgUtils_getImgBase64Code() error! and imgPath:" + imgPath + " Exception:" + e);
        }

        if (fis != null) {
            data = new byte[fis.available()];
            fis.read(data);
            fis.close();
        }

        BASE64Encoder encoder = new BASE64Encoder();
        if (data != null) {
            return encoder.encode(data);
        }

        return null;


    }


}
