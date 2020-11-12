package com.alibaba.easyexcel.test.demo.write;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * 基础数据类
 *
 * @author Jiaju Zhuang
 **/
@Data
@AllArgsConstructor
public class XHOrderSourceData {
    @ExcelProperty("订单来源")
    private String orderSource;
}
