package dev.drf.demo.poi.core.hssf;

import java.io.File;

public class HSSFPoiReaderBuilder {
    private File file;

    public static HSSFPoiReaderBuilder newBuilder() {
        return new HSSFPoiReaderBuilder();
    }

    public HSSFPoiReader build() {
        return new HSSFPoiReader(file);
    }

    public HSSFPoiReaderBuilder setFile(File file) {
        this.file = file;
        return this;
    }
}
