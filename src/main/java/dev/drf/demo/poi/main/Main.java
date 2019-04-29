package dev.drf.demo.poi.main;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Main {
    private static final Logger logger = LoggerFactory.getLogger(Main.class);
    public static void main(String[] args) {
        logger.debug("Log for: {}", Main.class.getCanonicalName());
    }
}
