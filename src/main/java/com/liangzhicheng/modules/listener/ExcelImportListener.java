package com.liangzhicheng.modules.listener;

import cn.hutool.core.util.ObjectUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.liangzhicheng.modules.annotation.ExcelValid;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ExcelImportListener extends AnalysisEventListener<Object> {

    @Override
    public void invoke(Object object, AnalysisContext analysisContext) {
        try{
            valid(object);
        }catch (Exception e){
            throw new ExcelAnalysisException(e.getMessage());
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

    }

    private void valid(Object object) throws Exception {
        List<Field> fieldList = new ArrayList<>();
        Class clazz = object.getClass();
        while(clazz != null){
            if(clazz.equals(Object.class)){
                clazz = clazz.getSuperclass();
                continue;
            }
            fieldList.addAll(new ArrayList<>(Arrays.asList(clazz.getDeclaredFields())));
            clazz = clazz.getSuperclass();
        }
        Field[] fields = new Field[fieldList.size()];
        fieldList.toArray(fields);
        for(Field field : fields){
            //设置可访问
            field.setAccessible(true);
            //属性的值
            Object fieldValue;
            try{
                fieldValue = field.get(object);
            }catch(IllegalAccessException e){
                throw new RuntimeException("导入参数校验失败");
            }
            if(field.isAnnotationPresent(ExcelValid.class) && ObjectUtil.isNull(fieldValue)){
                throw new IllegalAccessException(field.getAnnotation(ExcelValid.class).value());
            }
        }
    }

}