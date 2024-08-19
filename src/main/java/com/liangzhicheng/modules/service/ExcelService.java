package com.liangzhicheng.modules.service;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.builder.ExcelReaderSheetBuilder;
import com.liangzhicheng.modules.annotation.ExcelValid;
import com.liangzhicheng.modules.converter.LocalDateConverter;
import com.liangzhicheng.modules.converter.LocalDateTimeConverter;
import com.liangzhicheng.modules.converter.NumberLocalDateConverter;
import com.liangzhicheng.modules.listener.ExcelImportListener;
import lombok.Cleanup;
import org.springframework.core.io.ClassPathResource;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Service;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.List;

@Service
public class ExcelService {

    public void downloadTemplate_1() throws IOException {
        String fileName = "用户信息导入模板.xlsx";
        HttpServletResponse response = getResponse(fileName);
        ClassPathResource classPathResource = new ClassPathResource("static/template/excel/" + fileName);
        EasyExcel.write(response.getOutputStream())
                .withTemplate(classPathResource.getInputStream())
                .build()
                .finish();
    }

    public void downloadTemplate_2() {
        String fileName = "用户信息导入模板.xlsx";
        InputStream inputStream = super.getClass().getResourceAsStream("/static/template/excel/" + fileName);
        try{
            HttpServletResponse response = getResponse(fileName);
            @Cleanup BufferedInputStream in = new BufferedInputStream(inputStream);
            byte[] buf = new byte[1024];
            int len;
            @Cleanup OutputStream out = response.getOutputStream();
            while((len = in.read(buf)) > 0){
                out.write(buf, 0, len);
            }
        }catch(Exception e){
            System.out.println(e.getMessage());
        }
    }

    public <T> List<T> importExecute(MultipartFile file,
                                     Class<T> clazz,
                                     Integer headRowNumber) {
        List<T> excelList;
        try{
            ExcelReaderSheetBuilder excelReaderSheetBuilder =
                    EasyExcel.read(file.getInputStream()).head(clazz).sheet()
                            .registerConverter(new LocalDateConverter())
                            .registerConverter(new LocalDateTimeConverter())
                            .registerConverter(new NumberLocalDateConverter());
            if(clazz.isAnnotationPresent(ExcelValid.class)){
                excelReaderSheetBuilder.registerReadListener(new ExcelImportListener());
            }
            excelList = excelReaderSheetBuilder.headRowNumber(headRowNumber).doReadSync();
        }catch(Exception e){
            throw new RuntimeException(e.getMessage());
        }
        return excelList;
    }

    public <T> void exportExecute(String fileName,
                                  Class<T> clazz,
                                  List<T> list) {
        HttpServletResponse response = getResponse(fileName);
        try{
            EasyExcel.write(response.getOutputStream(), clazz)
                    .sheet("数据表")
                    .registerConverter(new LocalDateConverter())
                    .registerConverter(new LocalDateTimeConverter())
                    .useDefaultStyle(true)
                    .relativeHeadRowIndex(0)
                    .doWrite(list);
        }catch(IOException e){
            System.out.println("导出数据失败：" + e.getMessage());
        }
    }

    private HttpServletResponse getResponse(String fileName) {
        ServletRequestAttributes servletRequestAttributes = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        HttpServletResponse response = servletRequestAttributes.getResponse();
        response.setCharacterEncoding("UTF-8");
        response.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
        try{
            response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8"));
        }catch(UnsupportedEncodingException e){
            System.out.println(e.getMessage());
        }
        return response;
    }

}