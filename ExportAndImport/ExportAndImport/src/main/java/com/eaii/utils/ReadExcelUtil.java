package com.eaii.utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @Description: excel导入封装类
 * @author zxy
 * @date 2017年11月28日
 */
@Slf4j
public class ReadExcelUtil {

    public static <Q> List<Q> getExcelAsFile(String file,Class<Q> cls) throws FileNotFoundException, IOException, InvalidFormatException {
        List<Q> list = new ArrayList<Q>();
        Field[] fields = cls.getDeclaredFields();
        ArrayList<String> headList = new ArrayList<String>();
        InputStream ins = null;
        Workbook wb = null;
        ins=new FileInputStream(new File(file));
        //得到Excel工作簿对象
        wb = WorkbookFactory.create(ins);
        ins.close();
        int sheetNum = wb.getNumberOfSheets();
        boolean falg = false;
        for (int i = 0; i < sheetNum; i++) {
            //得到Excel工作表对象
            Sheet sheet = wb.getSheetAt(i);
            //总行数
            int trLength = sheet.getLastRowNum();
            //数据行读取
            for (int j = 1; j < trLength+1; j++) {
                try {
                    Q shixin = cls.newInstance();
                    Row row = sheet.getRow(j);
                    falg = false;
                    for (Field f : fields) {
                        ExcelField field = f.getAnnotation(ExcelField.class);
                        if (field != null){
                            //标题行
                            Row rowt = sheet.getRow(0);
                            for (int k = 0; k < rowt.getLastCellNum(); k++) {
                                Cell cellt = rowt.getCell(k);
                                int cellTypet = cellt.getCellType();
                                if(cellTypet==1&&cellt.getStringCellValue().equals(field.title())){
                                    falg = true;
                                    Cell cell = row.getCell(k);
                                    if(cell==null){
                                        falg = false;
                                        break;
                                    }
                                    int cellType = cell.getCellType();
                                    String typeName = f.getGenericType().getTypeName();
                                    //得到单元格类型
                                    switch (cellType){
                                        case 0:
                                            //CELL_TYPE_NUMERIC
                                            f.setAccessible(true);
                                            if("java.lang.String".equals(typeName)){
                                                if(cellType == HSSFCell.CELL_TYPE_NUMERIC){
                                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                                    f.set(shixin,cell.toString());
                                                }
                                            }else if("java.lang.Long".equals(typeName)){
                                                if(cellType == HSSFCell.CELL_TYPE_NUMERIC){
                                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                                    f.set(shixin,Long.parseLong(cell.toString()));
                                                }
                                            }else if("java.lang.Integer".equals(typeName)){
                                                f.set(shixin,new Double(cell.getNumericCellValue()).intValue());
                                            }else if("java.lang.Double".equals(typeName)){
                                                f.set(shixin,cell.getNumericCellValue());
                                            }
                                            break;
                                        case 1:
                                            //CELL_TYPE_STRING
                                            f.setAccessible(true);
                                            if("java.lang.String".equals(typeName)){
                                                f.set(shixin,cell.getStringCellValue());
                                            }else if("java.lang.Long".equals(typeName)){
                                                if(!"".equals(cell.getStringCellValue().trim())){
                                                    f.set(shixin,Long.parseLong(cell.getStringCellValue()));
                                                }else{
                                                    falg = false;
                                                }
                                            }else if("java.lang.Integer".equals(typeName)){
                                                if(!"".equals(cell.getStringCellValue().trim())){
                                                    f.set(shixin,Integer.parseInt(cell.getStringCellValue()));
                                                }else{
                                                    falg = false;
                                                }
                                            }else if("java.lang.Double".equals(typeName)){
                                                if(!"".equals(cell.getStringCellValue().trim())){
                                                    f.set(shixin,Double.parseDouble(cell.getStringCellValue()));
                                                }else{
                                                    falg = false;
                                                }
                                            }
                                            break;
                                        case 2:
                                            //CELL_TYPE_FORMULA
                                            //cell.getCellFormula();
                                            //System.out.println(cell.getCellFormula());
                                            break;
                                        case 3:
                                            //CELL_TYPE_BLANK
                                            break;
                                        case 4:
                                            //CELL_TYPE_BOOLEAN
                                            f.setAccessible(true);
                                            if("java.lang.Boolean".equals(typeName)){
                                                f.set(shixin,cell.getBooleanCellValue());
                                            }
                                            //cell.getBooleanCellValue();
                                            //System.out.println(cell.getBooleanCellValue());
                                            break;
                                        case 5:
                                            //CELL_TYPE_ERROR
                                            //cell.getErrorCellValue();
                                            //System.out.println(cell.getErrorCellValue());
                                            break;
                                        default:
                                            break;
                                    }
                                }
                            }
                        }
                    }
                    if(falg){
                        list.add(shixin);
                    }
                } catch (InstantiationException e) {
                    log.info("第"+j+"行出错");
                    log.info(e.getMessage());
                } catch (IllegalAccessException e) {
                    log.info("第"+j+"行出错");
                    log.info(e.getMessage());
                }
            }
        }
        return list;
    }
}
