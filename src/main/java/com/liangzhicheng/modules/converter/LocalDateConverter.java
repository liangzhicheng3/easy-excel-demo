package com.liangzhicheng.modules.converter;

import cn.hutool.core.date.DatePattern;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class LocalDateConverter extends EasyExcelBaseConverter<LocalDate> {

    public LocalDateConverter() {
        super(LocalDate.class);
    }

    @Override
    public LocalDate convertToJavaData(CellData cellData,
                                       ExcelContentProperty excelContentProperty,
                                       GlobalConfiguration globalConfiguration) throws Exception {
        if(cellData == null){
            return null;
        }
        LocalDate localDate = null;
        switch(cellData.getType()){
            case STRING:
                String stringValue = cellData.getStringValue();
                if(stringValue.contains("-")){
                    try{
                        localDate = LocalDate.parse(stringValue, DateTimeFormatter.ofPattern(DatePattern.NORM_DATE_PATTERN));
                    }catch(Exception e){
                        System.out.println("字符串日期格式转换LocalDate失败");
                    }
                }else if(stringValue.contains("/")){
                    String[] patterns = {"yyyy/M/d", "yyyy/M/dd", "yyyy/MM/dd"};
                    for(String pattern : patterns){
                        try{
                            localDate = LocalDate.parse(stringValue, DateTimeFormatter.ofPattern(pattern));
                        }catch(Exception e){
                            System.out.println("字符串日期格式转换LocalDate失败");
                        }
                    }
                }
                break;
        }
        return localDate;
    }

}