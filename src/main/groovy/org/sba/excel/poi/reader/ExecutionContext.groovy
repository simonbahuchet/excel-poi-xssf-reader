package org.sba.excel.poi.reader

import groovy.transform.PackageScope

class ExecutionContext {

    /** Map containing all the data related to reading: iteration, chunkItems being read, errors */
    private Map<String, Object> coreStore = [:]

    /** Anything the client application needs to store */
    Map<String, Object> store = [:]

    /** List of errors that occurred during the whole parsing */
    List<ExcelFileReaderError> errors = []

    /** Iteration counter */
    int iteration = 0

    /** List of chunkItems read during the iteration */
    List chunkItems = []

    /** List of sheet indices. Sheet that cannot be processed and which rows we skip */
    List<Integer> unprocessableSheets = []

    /** Current sheet, ie. sheet being read */
    Integer sheetIndex
    String sheetName

    /**
     * Default constructor
     */
    ExecutionContext() {
        init()
    }

    /**
     * Store the variable. Possibly overwrite it if it already existed
     */
    void put(String key, Object value) {
        store[key] = value
    }

    /**
     * Store the variable. Possibly overwrite it if it already existed
     */
    void add(Map<String, Object> params) {
        store += params
    }

    /**
     * Returned the stored variable or null if there's none for that key
     * @param key
     * @return
     */
    Object get(String key) {
        return store[key]
    }

    /**
     * Remove the value from the store
     * @param key
     */
    void remove(String key) {
        store.remove(key)
    }

    /**
     * Sdt the chunkItems being read
     * @param items
     */
    void addChunkItem(Object item) {
        chunkItems << item
        incIteration()
    }

    /**
     * Set the list of chunk items
     * @param items
     */
    @PackageScope
    void setChunkItems(List items) {
        chunkItems = items
    }

    /**
     * Set the iteration number
     * @param iteration
     */
    @PackageScope
    void setIteration(int iteration) {
        iteration = iteration
    }

    /**
     * Set the iteration number
     * @param iteration
     */
    Integer incIteration() {
        iteration++
        return iteration
    }

    /**
     * @return the list of errors that occurred while reading the chunk
     */
    List<ExcelFileReaderError> getErrors() {
        return errors
    }

    /**
     * @return true if there's at least one error
     */
    boolean hasErrors() {
        !!getErrors()
    }

    /**
     * Set the list of errors
     * @param errors
     */
    @PackageScope
    void setErrors(List errors) {
        this.errors = errors
    }

    /**
     * Set the sheet name
     */
    @PackageScope
    void setSheetName(String name) {
        sheetName = name
    }

    /**
     * Set the sheet index
     */
    @PackageScope
    void setSheetIndex(int index) {
        sheetIndex = index
    }

    /**
     * Add the unprocessable sheet number
     */
    void markSheetAsUnprocessable(Integer sheetNum) {
        unprocessableSheets << sheetNum
    }

    /**
     * Reset the list of unprocessable sheet numbers
     */
    void resetUnprocessableSheets() {
        unprocessableSheets.clear()
    }

    /**
     * @return the unprocessable sheets
     */
    boolean isUnprocessableSheet(Integer sheetNum) {
        return sheetNum in unprocessableSheets
    }

    /**
     * @return true if the current sheet is unprocessable
     */
    boolean isUnprocessableSheet() {
        return isUnprocessableSheet(this.sheetIndex)
    }

    /**
     * Reset the core technical attributes: errors, iteration, chunkItems
     */
    void init() {
        // Init the errors that might occur while reading the document
        errors.clear()

        // Init the iteration / row counter
        iteration = 0

        // Init the list of chunkItems composing the chunk
        chunkItems.clear()

        // re-init the execution context variables we set before reading
        unprocessableSheets.clear()
    }

    void closeChunk() {
        iteration = 0
        chunkItems.clear()
    }
}