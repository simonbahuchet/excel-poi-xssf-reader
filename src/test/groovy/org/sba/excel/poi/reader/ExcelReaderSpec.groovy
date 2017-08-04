package org.sba.excel.poi.reader

import groovy.util.logging.Slf4j
import spock.lang.Specification

@Slf4j
class ExcelReaderSpec extends Specification {

    void "Read a sample file should call the handlers"() {

        given:
        File file = new File("build/resources/test/file/Sample-Person-Data.xlsx")

        TestReadCallback readCallback = new TestReadCallback()
        TestSheetCallback sheetCallback = new TestSheetCallback()
        TestRowContentCallback rowContentCallback = new TestRowContentCallback()
        TestSheetSkippedCallback sheetSkippedCallback = new TestSheetSkippedCallback()
        TestRowSkippedCallback rowskippedCallback = new TestRowSkippedCallback()

        // Called every time a row is processed
        ExcelWorkSheetHandler workSheetHandler = new ExcelWorkSheetHandler(rowContentCallback)
        workSheetHandler.rowSkippedCallback = rowskippedCallback
        // What column to consider
        Map<String, String> columnHeaders = [:]
        columnHeaders << [("A"): "Person Id"]
        columnHeaders << [("B"): "Name"]
        //Ignore column C
        //columnHeaders << [("C"): "Height"]
        columnHeaders << [("D"): "Email Address"]
        columnHeaders << [("E"): "DOB"]
        columnHeaders << [("F"): "Salary"]
        workSheetHandler.columnHeaders = columnHeaders

        // Skip the first line that contains the headers
        workSheetHandler.rowsToSkip = [0]

        ExcelReader reader = new ExcelReader(file, workSheetHandler)
        reader.sheetCallback = sheetCallback
        reader.sheetSkippedCallback = sheetSkippedCallback
        reader.readCallback = readCallback

        //Ignore 2 of the 3 sheets
        reader.sheetsToSkip = ["Lot 1 Data", "Lot 3 Data"]

        when:
        reader.run([:])

        then:
        readCallback.numberOfStarts == 1
        readCallback.numberOfEnds == 1
        sheetCallback.numberOfStarts == 1
        sheetCallback.numberOfEnds == 1
        sheetSkippedCallback.numberOfCalls == 2
        rowContentCallback.numberOfCalls == 2
        rowskippedCallback.numberOfCalls == 1

        rowContentCallback.rowMaps.size() == 2

        rowContentCallback.rowMaps[1]["Person Id"] == "20001"
        rowContentCallback.rowMaps[1]["Name"] == "Jacob"
        rowContentCallback.rowMaps[1]["Height"] == null
        rowContentCallback.rowMaps[1]["Email Address"] == "jacob@example.example"
//        rowContentCallback.rowMaps[1]["DOB"] == "1-janv.-1970"
        rowContentCallback.rowMaps[1]["Salary"] == "10500"

        rowContentCallback.rowMaps[2]
    }
}