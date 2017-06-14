package org.sba.excel.poi.reader

import org.sba.excel.poi.reader.callback.ExcelRowSkippedCallback
import groovy.util.logging.Slf4j

@Slf4j
class TestRowSkippedCallback implements ExcelRowSkippedCallback {

    int numberOfCalls = 0

    @Override
    void skip(int rowNum) throws Exception {
        numberOfCalls++
    }
}
