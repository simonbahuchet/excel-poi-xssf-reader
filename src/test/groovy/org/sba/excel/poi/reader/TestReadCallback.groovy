package org.sba.excel.poi.reader

import org.sba.excel.poi.reader.callback.ExcelReadCallback
import groovy.util.logging.Slf4j

@Slf4j
class TestReadCallback implements ExcelReadCallback {

    int numberOfStarts = 0
    int numberOfEnds = 0

    @Override
    void beforeReading() throws Exception {
        log.debug "Start reading the document"
        numberOfStarts++
    }

    @Override
    void afterReading() throws Exception {
        log.debug "End reading the document"
        numberOfEnds++
    }
}
