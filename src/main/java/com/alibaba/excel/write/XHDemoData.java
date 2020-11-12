package com.alibaba.excel.write;

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
public class XHDemoData {
    @ExcelProperty("品牌")
    private String brandName;
    @ExcelProperty("门店")
    private String storeName;
    @ExcelProperty("营业日期")
    private String  businessDate;
    @ExcelProperty("订单来源")
    private String orderSource;
}
