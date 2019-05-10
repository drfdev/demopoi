package dev.drf.demo.poi.core.xssf;

import dev.drf.demo.poi.core.PoiListener;
import dev.drf.demo.poi.core.PoiReader;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Objects;

class XSSFPoiReader implements PoiReader {
    private OPCPackage pkg;
    XSSFPoiReader(File file) throws InvalidFormatException {
        this.pkg = OPCPackage.open(file);
    }

    enum ProcessType {
        FULL_FILE, SHEET_BY_NUMBER, SHEET_BY_NAME
    }

    @Override
    public void processFile(PoiListener listener)
            throws IOException, SAXException, OpenXML4JException, ParserConfigurationException {
        /*ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
        XSSFReader reader = new XSSFReader(pkg);

        StylesTable styles = reader.getStylesTable();
        XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) reader.getSheetsData();

        while (sheetIterator.hasNext()) {
            try (InputStream stream = sheetIterator.next();) {
                // String sheetName = iter.getSheetName();

                processingSheet(styles, strings, stream, listener);
            }
        }*/
        processWithEquals(listener, null, 0, ProcessType.FULL_FILE);
    }

    @Override
    public void processSheet(int sheetNum, PoiListener listener)
            throws IOException, SAXException, OpenXML4JException, ParserConfigurationException {
        /*ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
        XSSFReader reader = new XSSFReader(pkg);

        StylesTable styles = reader.getStylesTable();
        XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) reader.getSheetsData();

        int num = 0;
        while (sheetIterator.hasNext()) {
            try (InputStream stream = sheetIterator.next();) {
                if (num == sheetNum) {
                    processingSheet(styles, strings, stream, listener);
                    break;
                }
            }
            num ++;
        }*/
        processWithEquals(listener, null, sheetNum, ProcessType.SHEET_BY_NUMBER);
    }

    @Override
    public void processSheet(String sheetName, PoiListener listener)
            throws IOException, SAXException, OpenXML4JException, ParserConfigurationException {
        /*ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
        XSSFReader reader = new XSSFReader(pkg);

        StylesTable styles = reader.getStylesTable();
        XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) reader.getSheetsData();

        while (sheetIterator.hasNext()) {
            try (InputStream stream = sheetIterator.next();) {
                String thisSheetName = sheetIterator.getSheetName();
                if (Objects.equals(thisSheetName, sheetName)) {
                    processingSheet(styles, strings, stream, listener);
                    break;
                }
            }
        }*/
        processWithEquals(listener, sheetName, 0, ProcessType.SHEET_BY_NAME);
    }

    private void processWithEquals(PoiListener listener,
                                   String sheetName,
                                   int sheetNum,
                                   ProcessType processType)
            throws IOException, SAXException, OpenXML4JException, ParserConfigurationException {
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
        XSSFReader reader = new XSSFReader(pkg);

        StylesTable styles = reader.getStylesTable();
        XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) reader.getSheetsData();

        int num = 0;
        while (sheetIterator.hasNext()) {
            try (InputStream stream = sheetIterator.next()) {
                if (processType == ProcessType.SHEET_BY_NAME) {
                    String thisSheetName = sheetIterator.getSheetName();
                    if (Objects.equals(thisSheetName, sheetName)) {
                        processingSheet(styles, strings, stream, listener);
                        break;
                    }
                } else if (processType == ProcessType.SHEET_BY_NUMBER) {
                    if (num == sheetNum) {
                        processingSheet(styles, strings, stream, listener);
                        break;
                    }
                } else if (processType == ProcessType.FULL_FILE) {
                    processingSheet(styles, strings, stream, listener);
                }
                num ++;
            }
        }
    }

    private void processingSheet(StylesTable styles,
                                 ReadOnlySharedStringsTable strings,
                                 InputStream sheetInputStream,
                                 PoiListener listener)
            throws ParserConfigurationException, SAXException, IOException {
        InputSource sheetSource = new InputSource(sheetInputStream);

        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        SAXParser saxParser = saxFactory.newSAXParser();
        XMLReader sheetParser = saxParser.getXMLReader();

        ContentHandler handler = new XSSFSheetHandler(styles,
                strings,
                listener);

        sheetParser.setContentHandler(handler);
        sheetParser.parse(sheetSource);
    }

    @Override
    public void close() throws Exception {
        this.pkg.close();
    }
}
