package com.adrninistrator.texttoexcel.common;

import cn.idev.excel.write.handler.SheetWriteHandler;
import cn.idev.excel.write.metadata.holder.WriteSheetHolder;
import cn.idev.excel.write.metadata.holder.WriteWorkbookHolder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @author adrninistrator
 * @date 2025/8/16
 * @description:
 */
public class CustomSheetHandler implements SheetWriteHandler {
    private final int[] colWidths;
    private final int colCount;

    public CustomSheetHandler(int[] colWidths, int colCount) {
        this.colWidths = colWidths;
        this.colCount = colCount;
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder,
                                 WriteSheetHolder writeSheetHolder) {
        Sheet sheet = writeSheetHolder.getSheet();

        // 设置列宽
        for (int i = 0; i < colCount; i++) {
            sheet.setColumnWidth(i, colWidths[i]);
        }

        // 冻结首行
        sheet.createFreezePane(0, 1);

        // 启用筛选（从A1到最后一列）
        sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, colCount - 1));
    }
}
