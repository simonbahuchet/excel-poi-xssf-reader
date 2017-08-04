package org.sba.excel.poi.reader.callback

import groovy.util.logging.Slf4j
import org.sba.excel.poi.reader.ExecutionContext
import org.sba.excel.poi.reader.ExecutionContextAware

@Slf4j
class DefaultExcelSheetCallback implements ExcelSheetCallback, ExecutionContextAware {

    ExecutionContext executionContext

    @Override
    void setExecutionContext(ExecutionContext executionContext) {
        this.executionContext = executionContext
    }

    @Override
    void startSheet(int sheetIndex, String sheetName) {
        executionContext.sheetName = sheetName
        executionContext.sheetIndex = sheetIndex
    }

    @Override
    void endSheet(int sheetNum, String sheetName) {
        executionContext.sheetIndex = -1
        executionContext.sheetName = null
    }
}
