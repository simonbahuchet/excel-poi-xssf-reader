package org.sba.excel.poi.reader

/**
 * Exception that occurred during the parsing of an excel file
 */
class ExcelFileParseException extends Exception {

    ExcelFileParseException() {
        super
    }

    ExcelFileParseException(String message) {
        super(message)
    }

    ExcelFileParseException(String message, Throwable cause) {
        super(message, cause)
    }
}
