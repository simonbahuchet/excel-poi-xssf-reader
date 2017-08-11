package org.sba.excel.poi.reader.callback

import groovy.util.logging.Slf4j
import org.sba.excel.poi.reader.ExcelFileReaderError
import org.sba.excel.poi.reader.ExcelItemWriter
import org.sba.excel.poi.reader.ExecutionContext
import org.sba.excel.poi.reader.ExecutionContextAware

@Slf4j
class DefaultExcelReadCallback<T> implements ExcelReadCallback, ExecutionContextAware {

    /**
     * Store the core variables (chunkItems being read, errors, unprocessable sheets..) and the client app variables as well
     */
    ExecutionContext executionContext

    /**
     * Writer used to handle the chunkItems being read
     */
    ExcelItemWriter<T> writer

    @Override
    void setExecutionContext(ExecutionContext executionContext) {
        this.executionContext = executionContext
    }

    @Override
    void beforeReading() {
        // Reset the core variables: iteration, chunkItems, errors, un-processable sheets
        executionContext.init()
    }

    @Override
    void afterReading() {
        this.writeLastItems()
        executionContext.closeChunk()
    }

    /**
     * Write the chunkItems that might are left ; the chunk may not have been completed
     */
    void writeLastItems() {

        // Load the chunkItems stored during the read run
        List<T> items = executionContext.chunkItems

        if (items) {
            writer.writeLastItems(executionContext)
        }
    }
}
