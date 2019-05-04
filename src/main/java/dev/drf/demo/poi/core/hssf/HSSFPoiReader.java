package dev.drf.demo.poi.core.hssf;

import dev.drf.demo.poi.core.PoiListener;
import dev.drf.demo.poi.core.PoiReader;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Iterator;

class HSSFPoiReader implements PoiReader {
    private static final SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy");
    private static final DecimalFormat decimalFormat = new DecimalFormat("#.##");

    private HSSFWorkbook workbook;
    private POIFSFileSystem fs;

    HSSFPoiReader(File file) throws IOException {
        this.fs = new POIFSFileSystem(file, true);
        this.workbook = new HSSFWorkbook(fs);
    }

    @Override
    public void processFile(PoiListener listener) {
        Iterator<Sheet> sheetIter = workbook.sheetIterator();
        int sheetNum = 0;

        while (sheetIter.hasNext()) {
            Sheet sheet = sheetIter.next();
            listener.nextSheet(sheetNum ++);

            processingSheet(sheet, listener);
        }
    }

    @Override
    public void processSheet(int sheetNum, PoiListener listener) {
        Sheet sheet = workbook.getSheetAt(sheetNum);
        listener.nextSheet(workbook.getSheetIndex(sheet));
        processingSheet(sheet, listener);
    }

    @Override
    public void processSheet(String sheetName, PoiListener listener) {
        Sheet sheet = workbook.getSheet(sheetName);
        listener.nextSheet(workbook.getSheetIndex(sheet));
        processingSheet(sheet, listener);
    }

    private void processingSheet(Sheet sheet, PoiListener listener) {
        Iterator<Row> rowIter = sheet.rowIterator();
        int rowNum = 0;

        while (rowIter.hasNext()) {
            Row row = rowIter.next();
            listener.nextRow(rowNum);

            Iterator<Cell> cells = row.cellIterator();
            int cellNum = 0;

            while (cells.hasNext()) {
                Cell cell = cells.next();

                String val = getCellValue(cell);
                listener.readCell(cellNum, val);
            }
        }
    }

    @Override
    public void close() throws Exception {
        this.workbook.close();
        this.fs.close();
    }

    private static String getCellValue(Cell cell) {
        // cell --> string
        String value = "";
        CellType cellType = cell.getCellTypeEnum();
        switch (cellType) {
            case BLANK:
            case ERROR: {
                value = "";
                break;
            }
            case BOOLEAN: {
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            }
            case FORMULA: {
                try{
                    value = String.valueOf(cell.getNumericCellValue());
                } catch(Exception ex){
                    value = cell.getCellFormula();
                }
                break;
            }
            case NUMERIC: {
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    value = formatter.format(cell.getDateCellValue());
                } else {
                    value = decimalFormat.format(cell.getNumericCellValue());
                }
                break;
            }
            case STRING:
            default: {
                value = cell.getStringCellValue();
                break;
            }
        }
        return value; // not null
    }
}
