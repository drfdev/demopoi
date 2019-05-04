package dev.drf.demo.poi.core;

public interface PoiListener {

    void nextSheet(int sheetNum);
    void nextRow(int rowNum);
    void readCell(int cellNum, String value);
}
