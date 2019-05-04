package dev.drf.demo.poi.core;

public interface PoiReader extends AutoCloseable {

    void processFile(PoiListener listener);
    void processSheet(int sheetNum, PoiListener listener);
    void processSheet(String sheetName, PoiListener listener);
}
