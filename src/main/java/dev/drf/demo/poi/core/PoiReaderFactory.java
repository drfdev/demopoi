package dev.drf.demo.poi.core;

import dev.drf.demo.poi.core.err.PoiReaderException;
import dev.drf.demo.poi.core.hssf.HSSFPoiReaderBuilder;
import dev.drf.demo.poi.core.xssf.XSSFPoiReaderBuilder;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.file.Path;
import java.util.Objects;

public final class PoiReaderFactory {
    private static final Logger logger = LoggerFactory.getLogger(PoiReaderFactory.class);

    private static final String ENCRYPTED_PACKAGE = "EncryptedPackage";
    private static final int BUFFER_SIZE = 8;

    private PoiReaderFactory() {
        throw new AssertionError();
    }

    public enum TYPE {
        HSSF_WORKBOOK, // xls
        XSSF_WORKBOOK, // xlsx
        INVALID
    }

    public static PoiReader getInstance(File file) throws PoiReaderException, IOException, InvalidFormatException {
        if (!Objects.nonNull(file)) {
            throw new NullPointerException("File is null!");
        }

        TYPE type = fileType(file);

        switch (type) {
            case HSSF_WORKBOOK: {
//                return new HSSFPoiReader(file);
                return HSSFPoiReaderBuilder
                        .newBuilder()
                        .setFile(file)
                        .build();
            }
            case XSSF_WORKBOOK: {
//                return new XSSFPoiReader(file);
                return XSSFPoiReaderBuilder
                        .newBuilder()
                        .setFile(file)
                        .build();
            }
        }

        throw new PoiReaderException("Invalid workbook type");
    }

    public static PoiReader getInstance(Path path) throws PoiReaderException, IOException, InvalidFormatException {
        if (!Objects.nonNull(path)) {
            throw new NullPointerException("Path is null");
        }
        return getInstance(path.toFile());
    }

    private static TYPE fileType(File file) {
        try (InputStream inp = new FileInputStream(file)) {
            if (!(inp).markSupported()) {
                return getNotMarkSupportFileType(file);
            }
            return getType(inp);
        } catch (IOException e) {
            logger.error(e.getMessage(), e);
            return TYPE.INVALID;
        }
    }

    private static TYPE getNotMarkSupportFileType(File file) throws IOException {
//		try (InputStream inp = new PushbackInputStream(new FileInputStream(file), 8)) {
        try (InputStream inp = new BufferedInputStream(new FileInputStream(file), BUFFER_SIZE)) {
            return getType(inp);
        }
    }

    private static TYPE getType(InputStream inp) throws IOException {
        byte[] header8 = IOUtils.peekFirst8Bytes(inp);
        if (NPOIFSFileSystem.hasPOIFSHeader(header8)) {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(inp);
            return fileType(fs);
        } else if (DocumentFactoryHelper.hasOOXMLHeader(inp)) {
            return TYPE.XSSF_WORKBOOK;
        }
        return TYPE.INVALID;
    }

    private static TYPE fileType(NPOIFSFileSystem fs) {
        DirectoryNode root = fs.getRoot();
        if (root.hasEntry(ENCRYPTED_PACKAGE)) {
            return TYPE.XSSF_WORKBOOK;
        }
        return TYPE.HSSF_WORKBOOK;

    }
}
