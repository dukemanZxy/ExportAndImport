package com.eaii.utils;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 *
 * @Description: excel导出封装类
 * @author zxy
 * @date 2017年11月1日
 */
public class ExcelUtil
{
    //通用
    public static <Q> void writeExcel(HttpServletResponse response, String fileName, List<Q> list, Class<Q> cls) throws IOException, IllegalArgumentException, IllegalAccessException
    {
        HSSFWorkbook wb = new HSSFWorkbook();

        Field[] fields = cls.getDeclaredFields();
        ArrayList<String> headList = new ArrayList<String>();

        for (Field f : fields)
        {
            ExcelField field = f.getAnnotation(ExcelField.class);
            if (field != null)
            {
                headList.add(field.title());
            }
        }

        CellStyle style = getCellStyle(wb);
        List<List<Q>> listTemp = new ArrayList<List<Q>>();
        if(list!=null){
            int total = list.size();
            int cycelTotal = total / 65000;
            int tt = 0;
            for (int i = 0; i <= cycelTotal; i++) {
                if(((i+1)*65000)>total){
                    tt = total;
                }else{
                    tt = ((i+1)*65000);
                }
                listTemp.add(list.subList(i*65000,tt));
            }
        }
        for (int k = 0; k < listTemp.size(); k++) {
            Sheet sheet = wb.createSheet();
            /**
             * 设置Excel表的第一行即表头
             */
            Row row = sheet.createRow(0);
            for (int i = 0; i < headList.size(); i++)
            {
                Cell headCell = row.createCell(i);
                headCell.setCellType(Cell.CELL_TYPE_STRING);
                headCell.setCellStyle(style);//设置表头样式
                headCell.setCellValue(String.valueOf(headList.get(i)));
                sheet.autoSizeColumn((short) i);// 设置单元格自适应
                //sheet.setColumnWidth(0, 15 * 256);
            }
            if(listTemp.get(k)!=null){
                for (int i = 0; i < listTemp.get(k).size(); i++)
                {
                    Row rowdata = sheet.createRow(i + 1);//创建数据行
                    Q q = listTemp.get(k).get(i);
                    Field[] ff = q.getClass().getDeclaredFields();
                    int j = 0;
                    for (Field f : ff)
                    {
                        ExcelField field = f.getAnnotation(ExcelField.class);
                        if (field == null)
                        {
                            continue;
                        }
                        f.setAccessible(true);
                        Object obj = f.get(q);
                        Cell cell = rowdata.createCell(j);
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        // 当数字时
                        if (obj instanceof Integer) cell.setCellValue((Integer) obj);
                        // 当长整形
                        if (obj instanceof Long) cell.setCellValue((Long) obj);
                        // 当为字符串时
                        if (obj instanceof String) cell.setCellValue((String) obj);
                        // 当为布尔时
                        if (obj instanceof Boolean) cell.setCellValue((Boolean) obj);
                        // 当为时间时
                        if (obj instanceof Date) cell.setCellValue(MyUtils.getFormatDate("yyyy-MM-dd", (Date) obj));
                        // 当为时间时
                        if (obj instanceof Calendar) cell.setCellValue((Calendar) obj);
                        // 当为小数时
                        if (obj instanceof Double) cell.setCellValue((Double) obj);
                        //将序号替换为123456
                        if (j == 0) cell.setCellValue(i + 1);
                        j++;
                    }
                }
            }

        }
        response.setContentType( "application/octet-stream");
        response.setHeader("Content-Disposition", "attachment;filename=report.zip");
        OutputStream ouputStream = null;
        try
        {
            ouputStream = response.getOutputStream();
            ZipOutputStream zip = new ZipOutputStream(ouputStream);
            ZipEntry entry = new ZipEntry(MyUtils.urlEncode(fileName));
            zip.putNextEntry(entry);
            wb.write(zip);
            zip.flush();
            zip.close();
        }
        finally
        {
            ouputStream.flush();
            ouputStream.close();
        }
    }

    /**
     *
     * @Description:设置表头样式
     * @author zxy
     * @date 2017年11月1日
     */
    public static CellStyle getCellStyle(Workbook wb)
    {
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 12);//设置字体大小
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//加粗
        style.setFillForegroundColor(HSSFColor.LIME.index);// 设置背景色
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setAlignment(HSSFCellStyle.SOLID_FOREGROUND);//让单元格居中
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 左右居中
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 上下居中
        style.setWrapText(true);//设置自动换行
        style.setFont(font);
        return style;
    }

}