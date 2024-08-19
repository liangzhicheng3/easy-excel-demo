package com.liangzhicheng.modules.converter;

import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;

import java.time.LocalDate;
import java.time.temporal.ChronoUnit;

public class NumberLocalDateConverter extends EasyExcelBaseConverter<LocalDate> {

    public NumberLocalDateConverter() {
        super(LocalDate.class);
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.NUMBER;
    }

    @Override
    public LocalDate convertToJavaData(CellData cellData,
                                       ExcelContentProperty contentProperty,
                                       GlobalConfiguration globalConfiguration) throws Exception {
        if(cellData == null){
            return null;
        }
        LocalDate localDate = null;
        switch(cellData.getType()){
            case NUMBER:
                long days = ChronoUnit.DAYS.between(
                        LocalDate.of(1900, 1, 1),
                        LocalDate.of(1970, 1, 1));
                localDate = LocalDate.ofEpochDay(cellData.getNumberValue().longValue() - days - 2);
                break;
        }
        return localDate;
    }

}