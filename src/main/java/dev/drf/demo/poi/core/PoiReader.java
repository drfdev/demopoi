package dev.drf.demo.poi.core;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;

public interface PoiReader extends AutoCloseable {

    void processFile(PoiListener listener)
            throws IOException, SAXException, OpenXML4JException, ParserConfigurationException;
    void processSheet(int sheetNum, PoiListener listener);
    void processSheet(String sheetName, PoiListener listener);
}
