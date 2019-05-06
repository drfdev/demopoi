package dev.drf.demo.poi.core.xssf;

import dev.drf.demo.poi.core.PoiListener;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.helpers.DefaultHandler;

class XSSFSheetHandler extends DefaultHandler {
    XSSFSheetHandler(StylesTable styles,
                     ReadOnlySharedStringsTable strings,
                     PoiListener listener) {
        // TODO
    }
}
