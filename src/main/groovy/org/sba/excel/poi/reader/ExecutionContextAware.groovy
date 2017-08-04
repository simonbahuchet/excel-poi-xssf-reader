package org.sba.excel.poi.reader

/**
 * Characterize a class that needs to access the execution context
 */
interface ExecutionContextAware {

    /**
     * Set the execution context, before the reader starts reading the whole excel file
     *
     * @param executionContext
     */
    void setExecutionContext(ExecutionContext executionContext)
}