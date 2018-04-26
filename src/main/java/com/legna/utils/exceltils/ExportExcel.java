package com.legna.utils.exceltils;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
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
 * 较新版本Spring
 */
public class ExportExcel extends AbstractXlsView {
    private String head;
    private String[] titles;
    private Map<String,String> titleMap;

    /**
     *
     * @param head 表头
     * @param titleMap 表格项
     */
    public ExportExcel(String head,Map<String,String> titleMap) {
        this.head = head;
        this.titleMap = titleMap;
        Set<String> set = titleMap.keySet();
        String[] titles = new String[set.size()];
        titles = set.toArray(titles);
        this.titles = titles;
    }


    @Override
    protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request, HttpServletResponse response) throws Exception {
        //獲取數據
        List<Map<String, String>> list = (List<Map<String, String>>) model.get("excelList");
        //在workbook添加一個sheet
        HSSFSheet sheet = (HSSFSheet) workbook.createSheet(head);
        sheet.setDefaultColumnWidth(15);

        //设置表格标题
        HSSFRow row1 = sheet.createRow(0);
        HSSFCell cell1 = row1.createCell(0);
        cell1.setCellValue(head);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, titles.length));
        HSSFRow row2 = sheet.createRow(1);

        //遍历表格项
        HSSFCell cell = null;
        for (int i = 0; i < titles.length; i++) {
            cell = row2.createCell(i);
            cell.setCellValue(titleMap.get(titles[i]));
        }


        //添加数据
        for (int i = 0; i < list.size(); i++) {
            Map<String, String> map = list.get(i);
            HSSFRow row = sheet.createRow(i + 2);
            for (int j = 0; j < titles.length; j++) {
                //遍历各项，匹配到key值并填入
                String title = titles[j];
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

        /**
         * 此处可以加方法替换指定列，方法见下 transStatus()
         * 不推荐此方法，建议在dao层直接完成对数据的转换
         */

        workbook.write(ouputStream);
        ouputStream.flush();
        ouputStream.close();
    }


    /**
     * 类型转换
     * @param object
     * @return
     */
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

    /**
     * 根据需求 修改指定列
     * @param workbook
     * @param head
     * @param column
     * @param transMap
     * @return
     */
    public Workbook transStatus(Workbook workbook,String head,int column,Map<String,String> transMap) {
        Sheet sheet = workbook.getSheet(head);
        for(int i =2;i<sheet.getLastRowNum();i++){
            Cell cell = sheet.getRow(i).getCell(column);
            cell.setCellValue(transMap.get(cell.getStringCellValue()));
        }
        return workbook;
    }
}
