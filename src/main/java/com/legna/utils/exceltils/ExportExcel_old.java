package com.legna.utils.exceltils;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.web.servlet.view.document.AbstractXlsView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * 旧版本Spring
 * 本环境使用较新Spring，显示报错
 */
public class ExportExcel_old  extends AbstractExcelView {
    private String head;
    private String[] titles;
    private Map<String, String> titleMap;

    /**
     * @param head     表头
     * @param titleMap 表格项
     */
    public ExportExcel_old(String head, Map<String, String> titleMap) {
        this.head = head;
        this.titleMap = titleMap;
        Set<String> set = titleMap.keySet();
        String[] titles = new String[set.size()];
        titles = set.toArray(titles);
        this.titles = titles;
    }

    @Override
    public void buildExcelDocument(Map<String, Object> model,
                                   HSSFWorkbook workbook, HttpServletRequest request,
                                   HttpServletResponse response) throws Exception {
        //獲取數據
        List<Map<String, String>> list = (List<Map<String, String>>) model.get("excelList");
        //在workbook添加一個sheet
        HSSFSheet sheet = workbook.createSheet();
        sheet.setDefaultColumnWidth(15);
        HSSFCell cell = null;

        //设置表格标题
        HSSFRow row1 = sheet.createRow(0);
        HSSFCell cell1 = row1.createCell(0);
        cell1.setCellValue(head);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, titles.length));

        HSSFRow row2 = sheet.createRow(1);
        //遍歷標題
        for (int i = 0; i < titles.length; i++) {
            //獲取位置
            HSSFCell cell2 = row2.createCell(i);
            cell2.setCellValue(titleMap.get(titles[i]));
        }
        //數據寫出
        for (int i = 0; i < list.size(); i++) {
            //獲取每一個map
            Map<String, String> map = list.get(i);
            //一個map一行數據
            HSSFRow row = sheet.createRow(i + 2);
            for (int j = 0; j < titles.length; j++) {
                //遍歷標題，把key與標題匹配
                String title = titles[j];
                //判斷該內容存在mapzhong
                if (map.containsKey(title)) {
                    row.createCell(j).setCellValue(transType(map.get(title)));
                }
            }
        }
        //設置下載時客戶端Excel的名稱
        String filename = new SimpleDateFormat("yyyy-MM-dd").format(new Date()) + ".xls";
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-disposition", "attachment;filename=" + filename);
        OutputStream ouputStream = response.getOutputStream();
        workbook.write(ouputStream);
        ouputStream.flush();
        ouputStream.close();
    }

    public String transType(Object object) {
        if (object instanceof String) {
            String string = (String) object;
            return string;
        } else if (object instanceof Integer) {
            Integer integer = (Integer) object;
            return integer.toString();
        } else if (object instanceof Long) {
            Long aLong = (Long) object;
            return aLong.toString();
        } else if (object instanceof Short) {
            Short aShort = (Short) object;
            return aShort.toString();
        } else if (object instanceof Byte) {
            Byte aByte = (Byte) object;
            return aByte.toString();
        } else if (object instanceof Double) {
            Double aDouble = (Double) object;
            return aDouble.toString();
        } else if (object instanceof Float) {
            Float aFloat = (Float) object;
            return aFloat.toString();
        } else if (object instanceof Character) {
            Character character = (Character) object;
            return character.toString();
        } else if (object instanceof Boolean) {
            Boolean aBoolean = (Boolean) object;
            return aBoolean.toString();
        } else if (object instanceof Timestamp) {
            Timestamp timestamp = (Timestamp) object;
            return timestamp.toString();
        } else {
            return "";
        }
    }
}
