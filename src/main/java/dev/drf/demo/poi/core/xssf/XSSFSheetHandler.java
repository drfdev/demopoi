package dev.drf.demo.poi.core.xssf;

import dev.drf.demo.poi.core.PoiListener;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
//import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Objects;

class XSSFSheetHandler extends DefaultHandler {
    private StylesTable stylesTable;
    private ReadOnlySharedStringsTable sharedStringsTable;
    private PoiListener listener;

    private StringBuilder value;
    private XssfDataType nextDataType;

    private boolean vIsOpen;

    private int thisColumn;
    private int rowsCount;

    private short formatIndex;
    private String formatString;
    private final DataFormatter formatter;

    private static final SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");

    enum XssfDataType {
        BOOL,
        ERROR,
        FORMULA,
        INLINESTR,
        SSTINDEX,
        NUMBER
    }

    XSSFSheetHandler(StylesTable styles,
                     ReadOnlySharedStringsTable strings,
                     PoiListener listener) {
        this.stylesTable = styles;
        this.sharedStringsTable = strings;
        this.listener = listener;

        this.value = new StringBuilder();
        this.nextDataType = XssfDataType.NUMBER;

        this.formatter = new DataFormatter();

        this.thisColumn = -1;
        this.rowsCount = 0;
    }

    @Override
    public void startElement(String uri,
                             String localName,
                             String name,
                             Attributes attributes) /*throws SAXException*/ {
        if (Objects.equals("inlineStr", name)
                || Objects.equals("v", name)) {
            vIsOpen = true;
            // clear
            value.setLength(0);
        }
        // c => cell
        else if (Objects.equals("c", name)) {
            // r => reference
            String r = attributes.getValue("r");
            int firstDigit = -1;
            for (int c = 0; c < r.length(); ++c) {
                if (Character.isDigit(r.charAt(c))) {
                    firstDigit = c;
                    break;
                }
            }

            thisColumn = nameToColumn(r.substring(0, firstDigit));

            // default
            this.nextDataType = XssfDataType.NUMBER;
            this.formatIndex = -1;
            this.formatString = null;

            // t => cell type
            // s => cell style
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");

            if (Objects.equals("b", cellType)) {
                // b => bool
                nextDataType = XssfDataType.BOOL;
            } else if (Objects.equals("e", cellType)) {
                // e => error
                nextDataType = XssfDataType.ERROR;
            } else if (Objects.equals("inlineStr", cellType)) {
                // inline string
                nextDataType = XssfDataType.INLINESTR;
            } else if (Objects.equals("s", cellType)) {
                // sst index
                nextDataType = XssfDataType.SSTINDEX;
            } else if (Objects.equals("str", cellType)) {
                nextDataType = XssfDataType.FORMULA;
            } else if (cellStyleStr != null) {
                // It's a number, but almost certainly one
                // with a special style or format
                int styleIndex = Integer.parseInt(cellStyleStr);
                XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);

                this.formatIndex = style.getDataFormat();
                this.formatString = style.getDataFormatString();

                if (this.formatString == null) {
                    this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
                }
            }
        }
    }

    @Override
    public void endElement(String uri,
                           String localName,
                           String name) /*throws SAXException*/ {
        String thisStr; // = null;

        // v => contents of a cell
        if (Objects.equals("v", name)) {
            switch (nextDataType) {
                case BOOL: {
                    char first = value.charAt(0);
                    thisStr = first == '0' ? "false" : "true";
                    break;
                }
                case ERROR:
                case FORMULA:
                default: {
                    thisStr = value.toString();
                    break;
                }
                case INLINESTR: {
                    XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
                    thisStr = rtsi.toString();
                    break;
                }
                case SSTINDEX: {
                    String sstIndex = value.toString();
                    try {
                        int idx = Integer.parseInt(sstIndex);
                        XSSFRichTextString rtss = new XSSFRichTextString(
                                sharedStringsTable.getEntryAt(idx));
                        thisStr = rtss.toString();
                    } catch (NumberFormatException ex) {
                        thisStr = sstIndex;
                    }
                    break;
                }
                case NUMBER: {
                    String n = value.toString();
                    if (this.formatString != null) {
                        if (DateUtil.isADateFormat(this.formatIndex, this.formatString)) {
                            // date format
                            Date javaDate = DateUtil.getJavaDate(
                                    Double.parseDouble(n));
                            thisStr = dateFormat.format(javaDate);
                        } else {
                            thisStr = formatter.formatRawCellContents(Double.parseDouble(n),
                                    this.formatIndex,
                                    this.formatString);
                        }
                    } else {
                        thisStr = n;
                    }
                    break;
                }
            }

            // next cell value
            listener.readCell(thisColumn, thisStr);
        } else if (Objects.equals("row", name)) {
            // next row
            listener.nextRow(rowsCount ++);
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) /*throws SAXException*/ {
        if (vIsOpen) {
            value.append(ch, start, length);
        }
    }

    private int nameToColumn(String name) {
        int column = -1;
        for (int i = 0; i < name.length(); ++i) {
            int c = name.charAt(i);
            column = (column + 1) * 26 + c - 'A';
        }
        return column;
    }
}
