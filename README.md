# demopoi
Java POI demo

#### Libs:

**Java version:** 1.8

**POI version:** 3.17

**Testing:** jUnit 4.12, ~~mockito 1.9.5~~

**Logging:** logback 1.2.3, slf4g 1.7.26

#### Code example:

``` Java
PoiListener listener = ...;
Path filePath = Paths.get( ... );
try (PoiReader reader = PoiReaderFactory.getInstance(filePath)) {
    reader.processFile(listener);
} catch (Exception ex) {
    ...
}
```