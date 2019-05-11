package dev.drf.demo.poi;

import dev.drf.demo.poi.core.PoiReader;
import dev.drf.demo.poi.core.PoiReaderFactory;
import dev.drf.demo.poi.main.MyPoiListener;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.AfterClass;
import org.junit.Assert;
import org.junit.BeforeClass;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;
import java.util.Set;

public class MainTest {
    private static final Logger logger = LoggerFactory.getLogger(MainTest.class);
    private static final String TEST_FILE_NAME = String.join(File.separator,
            System.getProperty("user.home") ,
            "test.xlsx");
    private static final String SHEET_NAME = "First";

    @BeforeClass
    public static void beforeClass() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(SHEET_NAME);

        XSSFRow row0 = sheet.createRow(0);
        XSSFRow row1 = sheet.createRow(1);
        XSSFRow row2 = sheet.createRow(2);

        XSSFCell cell00 = row0.createCell(0, CellType.STRING);
        XSSFCell cell01 = row0.createCell(1, CellType.STRING);
        cell00.setCellValue("Name:");
        cell01.setCellValue("Value:");
        XSSFCell cell10 = row1.createCell(0, CellType.STRING);
        XSSFCell cell11 = row1.createCell(1, CellType.NUMERIC);
        cell10.setCellValue("Big");
        cell11.setCellValue(100);
        XSSFCell cell20 = row2.createCell(0, CellType.STRING);
        XSSFCell cell21 = row2.createCell(1, CellType.NUMERIC);
        cell20.setCellValue("Small");
        cell21.setCellValue(10);

        try (FileOutputStream fos = new FileOutputStream(TEST_FILE_NAME)) {
            workbook.write(fos);
        } catch (IOException ex) {
            logger.error(ex.toString(), ex);
        }
    }

    @Test
    public void readerTest() throws Exception {
        MyPoiListener myListener = new MyPoiListener();

        Path filePath = Paths.get(TEST_FILE_NAME);

        Assert.assertTrue(filePath.toFile().exists());

        try (PoiReader reader = PoiReaderFactory.getInstance(filePath)) {
            reader.processSheet(SHEET_NAME, myListener);
        } catch (Exception ex) {
            logger.error(ex.getMessage(), ex);
            throw ex;
        }

        Map<Integer, Map<Integer, String>> values = myListener.getSheetValues();

        Set<Integer> keys1 = values.keySet();
        Assert.assertTrue(keys1.contains(0));
        Assert.assertTrue(keys1.contains(1));
        Assert.assertTrue(keys1.contains(2));

        Map<Integer, String> row0 = values.get(0);
        Map<Integer, String> row1 = values.get(1);
        Map<Integer, String> row2 = values.get(2);

        Assert.assertNotNull(row0);
        Assert.assertNotNull(row1);
        Assert.assertNotNull(row2);

        Assert.assertEquals(row0.keySet().size(), 2);
        Assert.assertEquals(row1.keySet().size(), 2);
        Assert.assertEquals(row2.keySet().size(), 2);
    }

    @AfterClass
    public static void afterClass() {
        Path filePath = Paths.get(TEST_FILE_NAME);
        filePath.toFile().delete();
    }
}
