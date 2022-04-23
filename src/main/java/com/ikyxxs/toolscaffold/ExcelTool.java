package com.ikyxxs.toolscaffold;

import cn.hutool.core.io.FileUtil;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelUtil;

import java.util.Arrays;
import java.util.Collections;

/**
 * Excel工具
 */
public class ExcelTool {

    public static void main(String[] args) {
        // 处理后的Excel文件
        BigExcelWriter writer= ExcelUtil.getBigWriter("excel-after.xlsx");
        writer.write(Collections.singletonList(Arrays.asList("column1", "column2", "column3")));

        // 读取Excel文件，处理并写入新的文件
        ExcelUtil.readBySax(FileUtil.file("/Users/xxx/tool-scaffold/excel-before.xlsx"), 0,
                (sheetIndex, rowIndex, rows) -> {
//                    Console.log("[{}] [{}] {}", sheetIndex, rowIndex, rows);

                    if (rowIndex > 0) {
                        writer.write(Collections.singletonList(Arrays.asList(rows.get(0), rows.get(1), rows.get(0) + "" + rows.get(1))));
                    }
                });

        // 关闭writer，释放内存
        writer.close();
    }
}
