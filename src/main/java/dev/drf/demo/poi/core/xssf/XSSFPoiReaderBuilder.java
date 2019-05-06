package dev.drf.demo.poi.core.xssf;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;

public class XSSFPoiReaderBuilder {
    private File file;

    public static XSSFPoiReaderBuilder newBuilder() {
        return new XSSFPoiReaderBuilder();
    }

    public XSSFPoiReader build() throws InvalidFormatException {
        return new XSSFPoiReader(file);
    }

    public XSSFPoiReaderBuilder setFile(File file) {
        this.file = file;
        return this;
    }
}
