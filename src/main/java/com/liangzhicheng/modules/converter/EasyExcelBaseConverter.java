package com.liangzhicheng.modules.converter;

import cn.hutool.core.date.DatePattern;
import cn.hutool.core.date.DateUtil;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;

import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;

public abstract class EasyExcelBaseConverter<T> implements Converter<T> {

    private static final String DATE_PATTERN = DatePattern.NORM_DATE_PATTERN;

    private final Class<T> clazz;

    public EasyExcelBaseConverter(Class<T> clazz) {
        this.clazz = clazz;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    @Override
    public Class<T> supportJavaTypeKey() {
        return clazz;
    }

    @Override
    public CellData<String> convertToExcelData(T t, ExcelContentProperty excelContentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        if(t instanceof Date){
            return new CellData<>(DateUtil.format((Date) t, DATE_PATTERN));
        }
        if(t instanceof LocalDate){
            Instant instant = ((LocalDate) t).atStartOfDay().atZone(ZoneId.systemDefault()).toInstant();
            return new CellData<>(DateUtil.format(Date.from(instant), DATE_PATTERN));
        }
        if(t instanceof LocalDateTime){
            return new CellData<>(DateUtil.format(DateUtil.date((LocalDateTime) t), DATE_PATTERN));
        }
        return new CellData<>(t.toString());
    }

}