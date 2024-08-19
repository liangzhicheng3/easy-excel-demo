package com.liangzhicheng.modules.excel;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentRowHeight;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.liangzhicheng.modules.annotation.ExcelValid;
import lombok.Data;

@Data
@HeadRowHeight(20)
@ContentRowHeight(18)
public class UserExcel {

    @ColumnWidth(10)
    @ExcelValid("用户名称不能为空")
    @ExcelProperty("用户名称")
    private String userName;

    @ColumnWidth(10)
    @ExcelProperty("用户年龄")
    private String age;

}