package org.sba.excel.poi.reader

interface ExcelItemWriter<T> {

    /**
     * Write the chunkItems read during the iteration
     * @param items
     * @param errors
     * @return the number of insertions
     */
    int writeItems(List<T> items, List<ExcelFileReaderError> errors)

    /**
     * Write the last chunkItems. ie: the chunkItems that remain after all the iterations
     * @param items
     * @param errors
     * @return the number of insertions
     */
    int writeLastItems(List<T> items, List<ExcelFileReaderError> errors)
}
