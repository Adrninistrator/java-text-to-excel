package test.texttoexcel;

import com.adrninistrator.texttoexcel.entry.TextToExcelEntry;
import org.junit.Test;

/**
 * @author adrninistrator
 * @date 2025/8/16
 * @description:
 */
public class TestTextToExcel {

    public static final String TEXT_FILE_PATH = "src/test/resources/text.md";

    @Test
    public void testNormal() {
        TextToExcelEntry textToExcelEntry = new TextToExcelEntry(TEXT_FILE_PATH, 1440);
        textToExcelEntry.convertTextToExcel();
    }

    @Test
    public void testFontSize() {
        TextToExcelEntry textToExcelEntry = new TextToExcelEntry(TEXT_FILE_PATH, 1440, true);
        textToExcelEntry.setHeaderFontSize(20);
        textToExcelEntry.setContentFontSize(16);
        textToExcelEntry.convertTextToExcel();
    }
}
