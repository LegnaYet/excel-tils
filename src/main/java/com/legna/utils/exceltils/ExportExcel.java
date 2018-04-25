package com.legna.utils.exceltils;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.web.servlet.view.document.AbstractXlsView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class ExportExcel extends AbstractXlsView {
    private String head;
    private String[] titles;

    /**
     *
     * @param head 表头
     * @param titles 表格项
     */
    public ExportExcel(String head,String[] titles) {
        this.head=head;
        this.titles = titles;
    }


    @Override
    protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request, HttpServletResponse response) throws Exception {
        //獲取數據
        List<Map<String, String>> list = (List<Map<String, String>>) model.get("excelList");
        //在workbook添加一個sheet
        HSSFSheet sheet = (HSSFSheet) workbook.createSheet();
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
            cell.setCellValue(titles[i]);
        }


        //添加数据
        for (int i = 0; i < list.size(); i++) {
            Map<String, String> map = list.get(i);
            HSSFRow row = sheet.createRow(i + 2);
            for (int j = 0; j < titles.length; j++) {
                //遍历各项，匹配到key值并填入
                String title = titles[j];
                if (map.containsKey(title)) {
                    row.createCell(j).setCellValue(map.get(title));
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

    /**
     * 根据需求编写内容，替换数据库中字段的名称
     * @param title
     * @return
     */
    public String transTitleToZh(String title){
        switch (title){
            //JDK1.7之后case可以匹配字符串类型
            case "":
                return "";
        }
        return null;
    }
}
