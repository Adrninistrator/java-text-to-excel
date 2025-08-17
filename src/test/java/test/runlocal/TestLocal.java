package test.runlocal;

import com.adrninistrator.texttoexcel.entry.TextToExcelEntry;
import org.junit.Test;

import java.io.File;

/**
 * @author adrninistrator
 * @date 2025/8/17
 * @description:
 */
public class TestLocal {

    @Test
    public void test() {
        doText("D:\\gitee-dir\\pri-code\\java-text-to-excel\\src\\test\\resources\\localtextfile");
    }

    private void doText(String dirPath){
        File dir = new File(dirPath);
        for (File file : dir.listFiles()) {
            TextToExcelEntry textToExcelEntry = new TextToExcelEntry(file.getAbsolutePath(), 1280);
            textToExcelEntry.convertTextToExcel();
        }
    }
}
