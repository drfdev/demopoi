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
import java.util.Iterator;

class XSSFPoiReader implements PoiReader {
    private OPCPackage pkg;
    XSSFPoiReader(File file) throws InvalidFormatException {
        this.pkg = OPCPackage.open(file);
    }

    @Override
    public void processFile(PoiListener listener)
            throws IOException, SAXException, OpenXML4JException, ParserConfigurationException {
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
        XSSFReader reader = new XSSFReader(pkg);

        StylesTable styles = reader.getStylesTable();
        Iterator<InputStream> iter = reader.getSheetsData();
        // XSSFReader.SheetIterator iter

        while (iter.hasNext()) {
            InputStream stream = iter.next();
            // String sheetName = iter.getSheetName();

            processingSheet(styles, strings, stream, listener);
            stream.close();
        }
    }

    @Override
    public void processSheet(int sheetNum, PoiListener listener) {
        // TODO
    }

    @Override
    public void processSheet(String sheetName, PoiListener listener) {
        // TODO
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
