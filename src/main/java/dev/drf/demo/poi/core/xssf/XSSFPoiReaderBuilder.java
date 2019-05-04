package dev.drf.demo.poi.core.xssf;

import java.io.File;

public class XSSFPoiReaderBuilder {
    private File file;

    public static XSSFPoiReaderBuilder newBuilder() {
        return new XSSFPoiReaderBuilder();
    }

    public XSSFPoiReader build() {
        return new XSSFPoiReader(file);
    }

    public XSSFPoiReaderBuilder setFile(File file) {
        this.file = file;
        return this;
    }
}
