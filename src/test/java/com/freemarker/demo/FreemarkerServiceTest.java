package com.freemarker.demo;

import com.freemarker.demo.doc.DocumentFactory;
import com.freemarker.demo.doc.IDocument;
import com.freemarker.demo.doc.impl.FreemarkerService;
import com.freemarker.demo.util.ImgUtils;
import com.freemarker.demo.util.Json2xml;
import com.freemarker.demo.vo.Person;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import org.json.XML;
import org.junit.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class FreemarkerServiceTest {
    @Test
    public void exportDoc() throws Exception {
        IDocument IDocument = new FreemarkerService();

        Map<String, Object> dataMap = new HashMap<String, Object>();
        dataMap.put("name", "yacongliu");

        String msg = IDocument.exportDoc(dataMap, "doc/test.doc");


        System.out.println(msg);

    }

    @Test
    public void exportDocTable() throws Exception {
        Map<String, Object> dataMap = new HashMap<String, Object>();
        IDocument IDocument = new FreemarkerService("tem");

        dataMap.put("code", "001");
        dataMap.put("name", "张三" + ";<w:br />" + "李四");
        dataMap.put("number", "tbkg001");
        dataMap.put("phone", "125829321931");
        List list = new ArrayList();
        Map<String, String> map1 = new HashMap<String, String>(3);
        map1.put("jl", "2019-03 天津大学");


        Map<String, String> map2 = new HashMap<String, String>(3);
        map2.put("jl", "2015-03 南开大学");

        list.add(map1);
        list.add(map2);

        dataMap.put("columns", list);

        String msg = IDocument.exportDoc(dataMap, "doc/testTable.doc");

        System.out.println(msg);

    }

    @Test
    public void exportDocFactory() throws Exception {
        IDocument IDocument = DocumentFactory.produceFreemarker();

        Map<String, Object> dataMap = new HashMap<String, Object>(2);
        dataMap.put("name", "yacong_liu");

        String msg = IDocument.exportDoc(dataMap, "doc/test.doc");

        System.out.println(msg);

    }

    @Test
    public void getImgBase64CodeTest() throws IOException {
        String imgPath = "D:/tx.png";
        String imgBase64Code = ImgUtils.getImgBase64Code(imgPath);

        System.out.println(imgBase64Code);
    }

    @Test
    public void exportExcel(){
        try {
            File file = new File("doc/list.xlsx");
            FileInputStream fileInputStream = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheetAt(0);
            //总行数
            int rowNumTotal = sheet.getLastRowNum();

            List<Person> persons = new ArrayList<>();

            for (int i = 0; i < rowNumTotal; i++) {
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
            String msg = IDocument.exportDoc(dataMap, "doc/testTable.doc");

            System.out.println(msg);

        } catch (Exception e) {
            e.printStackTrace();
        }
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

    /**
     * xml 转 json
     *
     * @param xml
     * @return
     */
    public static String xml2json(String xml) throws Exception{
        JSONObject xmlJSONObj = XML.toJSONObject(xml);
        String jsonPrettyPrintString = xmlJSONObj.toString();
        return jsonPrettyPrintString;
    }

    public static void main(String[] args) {
        String xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n" +
                "\n" +
                "<ResInfo>\n" +
                "  <barcode/>\n" +
                "  <asset>\n" +
                "    <assetcardno>B</assetcardno>\n" +
                "    <comments/>\n" +
                "    <buydate/>\n" +
                "    <category/>\n" +
                "  </asset>\n" +
                "  <entity>\n" +
                "    <entityid>901061314</entityid>\n" +
                "    <entitycode>LTJHXYYCX01/XA-HWMACBTS529</entitycode>\n" +
                "    <entityname>蓝田局华胥电信营业厅CDMA基站/BTS529</entityname>\n" +
                "    <entityspec>BTS</entityspec>\n" +
                "    <vendorname>HuaWei</vendorname>\n" +
                "    <model>HUAWEI BTS3900</model>\n" +
                "    <installaddress/>\n" +
                "  </entity>\n" +
                "  <version/>\n" +
                "  <sectornum>3</sectornum>\n" +
                "  <rackname>HW401C</rackname>\n" +
                "  <containers>\n" +
                "    <container>\n" +
                "      <shelfhight>0.086</shelfhight>\n" +
                "      <cardinfos>\n" +
                "        <cardinfo>\n" +
                "          <cardname>290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框3槽(0)HECM</cardname>\n" +
                "        </cardinfo>\n" +
                "        <cardinfo>\n" +
                "          <cardname>290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框3槽(10)FAN</cardname>\n" +
                "        </cardinfo>\n" +
                "        <cardinfo>\n" +
                "          <cardname>290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框3槽(9)UPEU</cardname>\n" +
                "        </cardinfo>\n" +
                "        <cardinfo>\n" +
                "          <cardname>290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框3槽(7)CMPT</cardname>\n" +
                "        </cardinfo>\n" +
                "        <cardinfo>\n" +
                "          <cardname>290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框3槽(2)HCPM</cardname>\n" +
                "        </cardinfo>\n" +
                "      </cardinfos>\n" +
                "    </container>\n" +
                "    <container>\n" +
                "      <shelfhight>0.308</shelfhight>\n" +
                "      <cardinfos>\n" +
                "        <cardinfo>\n" +
                "          <cardname>290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框1槽(2)CRFU</cardname>\n" +
                "        </cardinfo>\n" +
                "        <cardinfo>\n" +
                "          <cardname>290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框1槽(4)CRFU</cardname>\n" +
                "        </cardinfo>\n" +
                "        <cardinfo>\n" +
                "          <cardname>290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框1槽(0)CRFU</cardname>\n" +
                "        </cardinfo>\n" +
                "      </cardinfos>\n" +
                "    </container>\n" +
                "    <container>\n" +
                "      <shelfhight>0.086</shelfhight>\n" +
                "    </container>\n" +
                "    <container>\n" +
                "      <shelfhight>0.044</shelfhight>\n" +
                "    </container>\n" +
                "  </containers>\n" +
                "  <ReturnResult>0</ReturnResult>\n" +
                "  <ReturnInfo>成功</ReturnInfo>\n" +
                "</ResInfo>";


        String json = "{\"ResInfo\":{\"sectornum\":3,\"ReturnResult\":0,\"containers\":{\"container\":[{\"shelfhight\":0.086,\"cardinfos\":{\"cardinfo\":[{\"cardname\":\"290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框3槽(0)HECM\"},{\"cardname\":\"290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框3槽(10)FAN\"},{\"cardname\":\"290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框3槽(9)UPEU\"},{\"cardname\":\"290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框3槽(7)CMPT\"},{\"cardname\":\"290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框3槽(2)HCPM\"}]}},{\"shelfhight\":0.308,\"cardinfos\":{\"cardinfo\":[{\"cardname\":\"290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框1槽(2)CRFU\"},{\"cardname\":\"290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框1槽(4)CRFU\"},{\"cardname\":\"290.蓝田局华胥电信营业厅CDMA基站/BTS529/架1列1/框1槽(0)CRFU\"}]}},{\"shelfhight\":0.086},{\"shelfhight\":0.044}]},\"rackname\":\"HW401C\",\"asset\":{\"comments\":\"\",\"buydate\":\"\",\"assetcardno\":\"B\",\"category\":\"\"},\"barcode\":\"\",\"version\":\"\",\"entity\":{\"entitycode\":\"LTJHXYYCX01/XA-HWMACBTS529\",\"installaddress\":\"\",\"entityspec\":\"BTS\",\"entityname\":\"蓝田局华胥电信营业厅CDMA基站/BTS529\",\"entityid\":901061314,\"model\":\"HUAWEI BTS3900\",\"vendorname\":\"HuaWei\"},\"ReturnInfo\":\"成功\"}}";

        String s = Json2xml.json2xml(json);
        System.out.println(s);
    }

}