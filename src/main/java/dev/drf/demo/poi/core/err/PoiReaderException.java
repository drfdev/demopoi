package dev.drf.demo.poi.core.err;

public class PoiReaderException extends Exception {
    public PoiReaderException() {
    }

    public PoiReaderException(String message) {
        super(message);
    }

    public PoiReaderException(String message, Throwable cause) {
        super(message, cause);
    }

    public PoiReaderException(Throwable cause) {
        super(cause);
    }

    public PoiReaderException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
