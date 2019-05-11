package dev.drf.demo.poi.main;

import dev.drf.demo.poi.core.PoiReader;
import dev.drf.demo.poi.core.PoiReaderFactory;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;

public class Main {
    private static final Logger logger = LoggerFactory.getLogger(Main.class);
    public static void main(String[] args) {
        logger.debug("Log for: {}", Main.class.getCanonicalName());

        final String fileName = String.join(File.separator,
                System.getProperty("user.home") ,
                "test.xlsx");
        final String sheetName = "First";

        logger.info("Create POI-file");
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(sheetName);

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

        try (FileOutputStream fos = new FileOutputStream(fileName)) {
            workbook.write(fos);
        } catch (IOException ex) {
            logger.error(ex.toString(), ex);
        }

        logger.info("Read POI-file");
        // PoiListener listener = ...;
        MyPoiListener myListener = new MyPoiListener();

        Path filePath = Paths.get(fileName);

        try (PoiReader reader = PoiReaderFactory.getInstance(filePath)) {
            reader.processSheet(sheetName, myListener);
        } catch (Exception ex) {
            logger.error(ex.getMessage(), ex);
        }

        logger.info("Sheet number:" + myListener.getSheetNum());
        logger.info("Sheet values:");
        Map<Integer, Map<Integer, String>> values = myListener.getSheetValues();
        for (Integer key1 : values.keySet()) {
            Map<Integer, String> rows = values.get(key1);
            for (Integer key2 : rows.keySet()) {
                String value = rows.get(key2);
                logger.debug(value);
            }
        }

        // delete file
        filePath.toFile().delete();
    }
}
