package com.freemarker.demo.controller;

import org.apache.http.Consts;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.mime.HttpMultipartMode;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.entity.mime.content.FileBody;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.ServletInputStream;
import javax.servlet.http.HttpServletRequest;
import java.io.*;
import java.util.HashMap;
import java.util.Map;

@Controller
@RequestMapping(value = "/httpUpload")
public class HttpUploadController {

    @RequestMapping(value = "/test")
    public @ResponseBody Map<String,Object> test(HttpServletRequest request){
        InputStream inputStream = null;
        try {
            byte [] bytes = new byte[1024];
            int rc = 0;
            inputStream = request.getInputStream();
            File file = new File("d:\\02.mp4");
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            while ((rc = inputStream.read(bytes,0, 1024))>0){
                fileOutputStream.write(bytes,0, rc);
            }
            fileOutputStream.flush();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        HashMap<String, Object> stringObjectHashMap = new HashMap<>();
        stringObjectHashMap.put("state",true);
        return stringObjectHashMap;
    }



    @RequestMapping(value = "/test2")
    public void test2(){
        CloseableHttpClient client = HttpClients.createDefault();

        HttpPost httpPost = new HttpPost("http://localhost:8081/httpUpload/test");

        InputStream inputStream = null;
        try {
            File file = new File("d:\\01.mp4");
            FileBody fileBody = new FileBody(file, ContentType.MULTIPART_FORM_DATA,"01.mp4");
            HttpEntity httpEntity = MultipartEntityBuilder
                    .create()
                    .setMode(HttpMultipartMode.BROWSER_COMPATIBLE)
                    .addPart("uploadFile", fileBody).build();
            httpPost.setEntity(httpEntity);

            client.execute(httpPost);
            HttpResponse response = client.execute(httpPost);
            httpEntity = response.getEntity();
            if(httpEntity != null){
                inputStream = httpEntity.getContent();
                //转换为字节输入流
                BufferedReader br = new BufferedReader(new InputStreamReader(inputStream, Consts.UTF_8));
                String body = null;
                while ((body = br.readLine()) != null) {
                    System.out.println(body);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
