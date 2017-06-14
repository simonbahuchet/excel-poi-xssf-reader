package org.sba.excel.poi.reader

import org.sba.excel.poi.reader.callback.ExcelSheetSkippedCallback
import groovy.util.logging.Slf4j

@Slf4j
class TestSheetSkippedCallback implements ExcelSheetSkippedCallback {

    int numberOfCalls = 0

    @Override
    void skip(int sheetNum, String sheetName) {
        log.debug "Skipping sheet $sheetName"
        numberOfCalls++
    }
}
