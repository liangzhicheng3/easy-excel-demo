package com.liangzhicheng.modules.controllers;

import com.liangzhicheng.modules.excel.UserExcel;
import com.liangzhicheng.modules.service.ExcelService;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;

@RequiredArgsConstructor
@RestController
public class ExcelController {

    private final ExcelService excelService;

    @PostMapping("/downloadTemplate_1")
    public void downloadTemplate_1() throws IOException {
        excelService.downloadTemplate_1();
    }

    @PostMapping("/downloadTemplate_2")
    public void downloadTemplate_2() {
        excelService.downloadTemplate_2();
    }

    @PostMapping("/importExecute")
    public void importExecute(@RequestParam("file") MultipartFile file) {
        excelService.importExecute(file, UserExcel.class, 1);
    }

    @PostMapping("/exportExecute")
    public void exportExecute() {
        //获取用户信息->userList，传到第三个参数
        excelService.exportExecute("用户信息报表", UserExcel.class, new ArrayList<>());
    }

}