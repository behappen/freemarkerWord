package com.freemarker.demo.controller;

import com.freemarker.demo.doc.IDocument;
import com.freemarker.demo.doc.impl.FreemarkerService;
import com.freemarker.demo.vo.Person;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Controller
public class DemoController {

    @RequestMapping(value = "/")
    public String index(){
        return "view/test";
    }

    @RequestMapping(value = "/exportExcel")
    public String exportExcel(HttpServletRequest request, HttpServletResponse response) throws Exception{
        MultipartHttpServletRequest multipartRequest = (MultipartHttpServletRequest) request;

        MultipartFile file = multipartRequest.getFile("filename");

        InputStream inputStream = file.getInputStream();

        ////new XSSFWorkbook(inputStream);
        Workbook workbook = getWorkbook(inputStream,file.getOriginalFilename());

        Sheet sheet = workbook.getSheetAt(0);
        //总行数
        int rowNumTotal = sheet.getLastRowNum();

        List<Person> persons = new ArrayList<>();

        for (int i = 0; i <= rowNumTotal; i++) {
            //获取行
            Row row = sheet.getRow(i);
            //总列数
            int coloumNumTotal = row.getPhysicalNumberOfCells();
            for (int j = 0; j < coloumNumTotal; j++) {
                Person person = new Person();
                Cell namCell = row.getCell(j);
                String name = namCell.getStringCellValue();
                person.setName(name.replace("/"," "));
                Cell birthdayCell = row.getCell(j+1);
                String birthday = birthdayCell.getStringCellValue();
                person.setBirthday(formatDate(birthday));
                Cell passportNumberCell = row.getCell(j + 2);
                String passportNumber = passportNumberCell.getStringCellValue();
                person.setPassportNumber(passportNumber);
                Cell passportDateCell = row.getCell(j + 3);
                String passportDate = passportDateCell.getStringCellValue();
                person.setPassportDate(formatDate(passportDate));
                persons.add(person);
                break;
            }

        }
        //加载模板
        Map<String, Object> dataMap = new HashMap<>();
        IDocument IDocument = new FreemarkerService("temp");
        //生成word文档
        dataMap.put("columns", persons);
        IDocument.exportDoc("自动生成word表", dataMap, response);
        return null;
    }

    /**
     * 判断文件格式
     *
     * @param inStr
     * @param fileName
     * @return
     * @throws Exception
     */
    public Workbook getWorkbook(InputStream inStr, String fileName) throws Exception {
        Workbook workbook = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if (".xls".equals(fileType)) {
            workbook = new HSSFWorkbook(inStr);
        } else if (".xlsx".equals(fileType)) {
            workbook = new XSSFWorkbook(inStr);
        } else {
            throw new Exception("请上传excel文件！");
        }
        return workbook;
    }


    public static String formatDate(String date){
        StringBuffer sb = new StringBuffer();
        String[] dateArray = date.split("/");
        //yyyy-MM-dd  -->  dd-MM-yyyy
        sb.append(dateArray[2].length() < 2 ? "0" + dateArray[2] : dateArray[2]).append("/");
        sb.append(dateArray[1].length() < 2 ? "0" + dateArray[1] : dateArray[1]).append("/");
        sb.append(dateArray[0]);
        return sb.toString();
    }
}
