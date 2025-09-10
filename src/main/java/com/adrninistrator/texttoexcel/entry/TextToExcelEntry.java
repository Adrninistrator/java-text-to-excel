package com.adrninistrator.texttoexcel.entry;

import cn.idev.excel.EasyExcel;
import cn.idev.excel.write.metadata.style.WriteCellStyle;
import cn.idev.excel.write.metadata.style.WriteFont;
import cn.idev.excel.write.style.HorizontalCellStyleStrategy;
import com.adrninistrator.texttoexcel.common.CustomSheetHandler;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedReader;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * @author adrninistrator
 * @date 2025/8/16
 * @description:
 */
public class TextToExcelEntry {

    private static final Logger logger = LoggerFactory.getLogger(TextToExcelEntry.class);

    public static final int MIN_COLUMN_WIDTH = 3 * 256;
    public static final int MAX_COLUMN_WIDTH = 255 * 256;

    // 需要转换为excel文件的文本文件路径
    private final String inputFilePath;

    // 生成的excel文件总宽度像素（近似值，每列宽度受最小与最大值限制）
    private final int totalWidth;

    // 是否在生成的excel文件名末尾拼接时间戳，以避免不允许打开同名excel文件的问题
    private final boolean appendTimestampInFileName;

    // 生成的excel文件路径
    private String outputExcelPath;

    // 文本文件用于分隔列的字符
    private String splitChar = "\t";

    // 首行字体大小
    private int headerFontSize = 11;
    // 非首行字体大小
    private int contentFontSize = 11;

    /**
     * @param inputFilePath 需要转换为excel文件的文本文件路径
     * @param totalWidth    生成的excel文件总宽度像素（近似值，每列宽度受最小与最大值限制）
     */
    public TextToExcelEntry(String inputFilePath, int totalWidth) {
        this(inputFilePath, totalWidth, false);
    }

    /**
     * @param inputFilePath             需要转换为excel文件的文本文件路径
     * @param totalWidth                生成的excel文件总宽度像素（近似值，每列宽度受最小与最大值限制）
     * @param appendTimestampInFileName 是否在生成的excel文件名末尾拼接时间戳，以避免不允许打开同名excel文件的问题
     */
    public TextToExcelEntry(String inputFilePath, int totalWidth, boolean appendTimestampInFileName) {
        this.inputFilePath = inputFilePath;
        this.totalWidth = totalWidth;
        this.appendTimestampInFileName = appendTimestampInFileName;
    }

    public boolean convertTextToExcel() {
        int headerColumnNum = -1;
        int lineNum = 0;
        // 1. 读取文本文件并解析数据
        List<String[]> allData = new ArrayList<>();
        try (BufferedReader br = new BufferedReader(new FileReader(inputFilePath))) {
            String line;
            while ((line = br.readLine()) != null) {
                lineNum++;
                String[] data = StringUtils.splitPreserveAllTokens(line, splitChar);
                if (headerColumnNum == -1) {
                    headerColumnNum = data.length;
                } else if (headerColumnNum != data.length) {
                    logger.error("{} 第 {} 行文件内容 [{}] 分隔后的列数 {} 与首行的列数 {} 不同", inputFilePath, lineNum, line, data.length, headerColumnNum);
                    return false;
                }
                allData.add(data);
            }
        } catch (Exception e) {
            logger.error("error ", e);
            return false;
        }

        if (allData.isEmpty()) {
            logger.warn("文件内容为空 {}", inputFilePath);
            return true;
        }
        // 2. 计算每列平均字节数
        int colCount = allData.get(0).length;
        double[] colAvgBytes = new double[colCount];

        for (int col = 0; col < colCount; col++) {
            double totalBytes = 0;
            int rowNumWithContent = 0;
            for (String[] row : allData) {
                if (col < row.length) {
                    int byteCount = calculateByteCount(row[col]);
                    totalBytes += byteCount;
                    if (byteCount > 0) {
                        rowNumWithContent++;
                    }
                }
            }
            colAvgBytes[col] = totalBytes / rowNumWithContent;
        }

        // 3. 计算列宽比例（基于总宽度）
        int[] colWidths = new int[colCount];
        double totalAvgBytes = 0;
        for (double bytes : colAvgBytes) {
            totalAvgBytes += bytes;
        }

        for (int i = 0; i < colCount; i++) {
            double ratio = colAvgBytes[i] / totalAvgBytes;
            // 计算列宽
            int width = (int) (totalWidth * ratio * 32);
            // 设置列宽，最大不超过255*256，最小不小于10
            width = Math.min(width, MAX_COLUMN_WIDTH);
            width = Math.max(width, MIN_COLUMN_WIDTH);
            colWidths[i] = width;
        }

        // 4. 准备写入Excel的数据（移除标题行）
        List<List<String>> dataRows = new ArrayList<>();
        for (int i = 1; i < allData.size(); i++) {
            List<String> rowData = new ArrayList<>();
            Collections.addAll(rowData, allData.get(i));
            dataRows.add(rowData);
        }

        // 5. 准备表头
        List<List<String>> headerList = new ArrayList<>();
        String[] headerData = allData.get(0);
        for (String header : headerData) {
            List<String> tmpList = new ArrayList<>();
            tmpList.add(header);
            headerList.add(tmpList);
        }

        // 6. 创建Excel文件
        String outputFilePath;
        if (StringUtils.isNotBlank(outputExcelPath)) {
            outputFilePath = outputExcelPath;
        } else if (appendTimestampInFileName) {
            outputFilePath = inputFilePath + "_" + System.currentTimeMillis() + ".xlsx";
        } else {
            outputFilePath = inputFilePath + ".xlsx";
        }
        logger.info("输入的文本文件路径 {} 生成excel文件路径 {}", inputFilePath, outputFilePath);

        HorizontalCellStyleStrategy horizontalCellStyleStrategy = genHorizontalCellStyleStrategy();
        // 7. 注册自定义处理器（处理列宽、冻结窗格、筛选器）
        EasyExcel.write(outputFilePath)
                .registerWriteHandler(new CustomSheetHandler(colWidths, colCount))
                .registerWriteHandler(horizontalCellStyleStrategy)
                .head(headerList)
                .sheet("Sheet1")
                .doWrite(dataRows);
        return true;
    }

    private HorizontalCellStyleStrategy genHorizontalCellStyleStrategy() {
        // 头的策略
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setFontHeightInPoints((short) headerFontSize);
        headWriteCellStyle.setWriteFont(headWriteFont);
        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        WriteFont contentWriteFont = new WriteFont();
        // 字体大小
        contentWriteFont.setFontHeightInPoints((short) contentFontSize);
        contentWriteCellStyle.setWriteFont(contentWriteFont);
        // 这个策略是 头是头的样式 内容是内容的样式 其他的策略可以自己实现
        return new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
    }

    // 计算字符串的字节长度（中文2字节，英文1字节）
    private int calculateByteCount(String str) {
        int count = 0;
        for (char c : str.toCharArray()) {
            count += (c >= 0x4E00 && c <= 0x9FA5) ? 2 : 1;
        }
        return count;
    }

    public void setOutputExcelPath(String outputExcelPath) {
        this.outputExcelPath = outputExcelPath;
    }

    public void setSplitChar(String splitChar) {
        this.splitChar = splitChar;
    }

    public void setHeaderFontSize(int headerFontSize) {
        this.headerFontSize = headerFontSize;
    }

    public void setContentFontSize(int contentFontSize) {
        this.contentFontSize = contentFontSize;
    }
}
