package com.liangzhicheng.modules.converter;

import cn.hutool.core.date.DatePattern;
import cn.hutool.core.date.DateUtil;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;

import java.time.LocalDateTime;

public class LocalDateTimeConverter extends EasyExcelBaseConverter<LocalDateTime> {

    public LocalDateTimeConverter() {
        super(LocalDateTime.class);
    }

    @Override
    public LocalDateTime convertToJavaData(CellData cellData,
                                           ExcelContentProperty contentProperty,
                                           GlobalConfiguration globalConfiguration) throws Exception {
        if(cellData == null){
            return null;
        }
        LocalDateTime localDateTime = null;
        switch(cellData.getType()){
            case STRING:
                String stringValue = cellData.getStringValue();
                if(stringValue.contains("-")){
                    try{
                        localDateTime = DateUtil.toLocalDateTime(DateUtil.parse(stringValue, DatePattern.NORM_DATE_PATTERN));
                    }catch(Exception e){
                        System.out.println("字符串日期格式转换LocalDateTime失败");
                    }
                }else if(stringValue.contains("/")){
                    try{
                        localDateTime = DateUtil.toLocalDateTime(DateUtil.parse(stringValue, "yyyy/MM/dd"));
                    }catch(Exception e){
                        System.out.println("字符串日期格式转换LocalDateTime失败");
                    }
                }
                break;
        }
        return localDateTime;
    }

}