package org.sba.excel.poi.reader

interface ExcelItemWriter<T> {

    /**
     * Get the items from the execution context (and access any other variable stored within the context) and write them
     * @param executionContext
     * @return
     */
    int writeItems(ExecutionContext executionContext)

    /**
     * Write the last executionContext.chunkItems. ie: the chunkItems that remain after all the iterations
     *
     * @param items
     * @param errors
     * @return the number of insertions
     */
    int writeLastItems(ExecutionContext executionContext)
}
