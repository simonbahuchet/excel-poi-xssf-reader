package org.sba.excel.poi.reader

import groovy.util.logging.Slf4j
import spock.lang.Specification

@Slf4j
class ExcelReaderSpec extends Specification {

    void "Read a sample file should call the handlers with configured column headers"() {

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

        //Ignore 3 of the 4 sheets
        reader.sheetsToSkip = ["Lot 1 Data", "Lot 3 Data", "Lot 4 Data"]

        when:
        reader.run([:])

        then:
        readCallback.numberOfStarts == 1
        readCallback.numberOfEnds == 1
        sheetCallback.numberOfStarts == 1
        sheetCallback.numberOfEnds == 1
        sheetSkippedCallback.numberOfCalls == 3
        rowContentCallback.numberOfCalls == 2
        rowskippedCallback.numberOfCalls == 1

        rowContentCallback.rowMaps.size() == 2

        Map<String, String> mapFieldValue = rowContentCallback.rowMaps[1]
        mapFieldValue.size()==5
        rowContentCallback.rowMaps[1]["Person Id"] == "20001"
        rowContentCallback.rowMaps[1]["Name"] == "Jacob"
        rowContentCallback.rowMaps[1]["Height"] == null
        rowContentCallback.rowMaps[1]["Email Address"] == "jacob@example.example"
//        rowContentCallback.rowMaps[1]["DOB"] == "1-janv.-1970"
        rowContentCallback.rowMaps[1]["Salary"] == "10500"

        rowContentCallback.rowMaps[2]
    }

    void "Read a sample fileInputStream should call the handlers with configured column headers"() {

        given:

        File file = new File("build/resources/test/file/Sample-Person-Data.xlsx")

        InputStream inputStream = new FileInputStream(file)

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

        ExcelReader reader = new ExcelReader(inputStream, workSheetHandler)
        reader.sheetCallback = sheetCallback
        reader.sheetSkippedCallback = sheetSkippedCallback
        reader.readCallback = readCallback

        //Ignore 2 of the 3 sheets
        reader.sheetsToSkip = ["Lot 1 Data", "Lot 3 Data", "Lot 4 Data"]

        when:
        reader.run([:])

        then:
        readCallback.numberOfStarts == 1
        readCallback.numberOfEnds == 1
        sheetCallback.numberOfStarts == 1
        sheetCallback.numberOfEnds == 1
        sheetSkippedCallback.numberOfCalls == 3
        rowContentCallback.numberOfCalls == 2
        rowskippedCallback.numberOfCalls == 1

        rowContentCallback.rowMaps.size() == 2

        Map<String, String> mapFieldValue = rowContentCallback.rowMaps[1]
        mapFieldValue.size()==5

        rowContentCallback.rowMaps[1]["Person Id"] == "20001"
        rowContentCallback.rowMaps[1]["Name"] == "Jacob"
        rowContentCallback.rowMaps[1]["Height"] == null
        rowContentCallback.rowMaps[1]["Email Address"] == "jacob@example.example"
//        rowContentCallback.rowMaps[1]["DOB"] == "1-janv.-1970"
        rowContentCallback.rowMaps[1]["Salary"] == "10500"

        rowContentCallback.rowMaps[2]
    }

    void "Read a sample file should call the handlers with default column headers"() {

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
        // We consider all columns in header mode

        // in header mode => the header row is "skipped" so don't specify this row as rowsToSkip
        workSheetHandler.rowsToSkip = []

        ExcelReader reader = new ExcelReader(file, workSheetHandler)
        reader.sheetCallback = sheetCallback
        reader.sheetSkippedCallback = sheetSkippedCallback
        reader.readCallback = readCallback

        //Ignore 2 of the 3 sheets
        reader.sheetsToSkip = ["Lot 1 Data", "Lot 3 Data", "Lot 4 Data"]

        when:
        reader.run([:])

        then:
        readCallback.numberOfStarts == 1
        readCallback.numberOfEnds == 1
        sheetCallback.numberOfStarts == 1
        sheetCallback.numberOfEnds == 1
        sheetSkippedCallback.numberOfCalls == 3
        rowContentCallback.numberOfCalls == 2
        rowskippedCallback.numberOfCalls == 0

        rowContentCallback.rowMaps.size() == 2

        Map<String, String> mapFieldValue = rowContentCallback.rowMaps[1]
        mapFieldValue.size()==6

        rowContentCallback.rowMaps[1]["Person Id"] == "20001"
        rowContentCallback.rowMaps[1]["Name"] == "Jacob"
        rowContentCallback.rowMaps[1]["Height"] == "5,20"
        rowContentCallback.rowMaps[1]["Email Address"] == "jacob@example.example"
        rowContentCallback.rowMaps[1]["DOB"] == "1-janv.-1970"
        rowContentCallback.rowMaps[1]["Salary"] == "10500"

        rowContentCallback.rowMaps[2]
    }


    void "Read a sample file should call the handlers with excel column reference as headers"() {

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
        // We consider all columns in header mode

        workSheetHandler.mappingMode = ExcelWorkSheetHandler.ColumnMapping.REFERENCE_AND_HEADER_ROW

        // We skip the header row cause we are in column reference mode
        workSheetHandler.rowsToSkip = []

        ExcelReader reader = new ExcelReader(file, workSheetHandler)
        reader.sheetCallback = sheetCallback
        reader.sheetSkippedCallback = sheetSkippedCallback
        reader.readCallback = readCallback

        //Ignore 2 of the 3 sheets
        reader.sheetsToSkip = ["Lot 1 Data", "Lot 3 Data", "Lot 4 Data"]

        when:
        reader.run([:])

        then:
        readCallback.numberOfStarts == 1
        readCallback.numberOfEnds == 1
        sheetCallback.numberOfStarts == 1
        sheetCallback.numberOfEnds == 1
        sheetSkippedCallback.numberOfCalls == 3
        rowContentCallback.numberOfCalls == 2
        rowskippedCallback.numberOfCalls == 0

        rowContentCallback.rowMaps.size() == 2

        Map<String, String> mapFieldValue = rowContentCallback.rowMaps[1]
        mapFieldValue.size()==6

        rowContentCallback.rowMaps[1]["A"] == "20001"
        rowContentCallback.rowMaps[1]["B"] == "Jacob"
        rowContentCallback.rowMaps[1]["C"] == "5,20"
        rowContentCallback.rowMaps[1]["D"] == "jacob@example.example"
        rowContentCallback.rowMaps[1]["E"] == "1-janv.-1970"
        rowContentCallback.rowMaps[1]["F"] == "10500"

        reader.sheetContentsHandler.columnHeaders["A"]== "Person Id"
        reader.sheetContentsHandler.columnHeaders["B"]== "Name"
        reader.sheetContentsHandler.columnHeaders["C"]== "Height"
        reader.sheetContentsHandler.columnHeaders["D"]== "Email Address"
        reader.sheetContentsHandler.columnHeaders["E"]== "DOB"
        reader.sheetContentsHandler.columnHeaders["F"]== "Salary"

        rowContentCallback.rowMaps[2]
    }

    void "Read a sample file should skip 2 first rows and call the handlers with excel column reference as headers"() {

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
        // We consider all columns in header mode

        workSheetHandler.mappingMode = ExcelWorkSheetHandler.ColumnMapping.REFERENCE_AND_HEADER_ROW

        // We skip 2 first rows
        workSheetHandler.rowsToSkip = [0,1]

        // We consider row 2 as header
        workSheetHandler.headerRow = 2

        ExcelReader reader = new ExcelReader(file, workSheetHandler)
        reader.sheetCallback = sheetCallback
        reader.sheetSkippedCallback = sheetSkippedCallback
        reader.readCallback = readCallback

        //Ignore sheet 2 & 3
        reader.sheetsToSkip = ["Lot 2 Data", "Lot 3 Data", "Lot 4 Data"]

        when:
        reader.run([:])

        then:
        readCallback.numberOfStarts == 1
        readCallback.numberOfEnds == 1
        sheetCallback.numberOfStarts == 1
        sheetCallback.numberOfEnds == 1
        sheetSkippedCallback.numberOfCalls == 3
        rowContentCallback.numberOfCalls == 5
        rowskippedCallback.numberOfCalls == 2

        rowContentCallback.rowMaps.size() == 5

        Map<String, String> mapFieldValue = rowContentCallback.rowMaps[3]
        mapFieldValue.size()==6

        rowContentCallback.rowMaps[3]["A"] == "10002"
        rowContentCallback.rowMaps[3]["B"] == "Emily"
        rowContentCallback.rowMaps[3]["C"] == "5,40"
        rowContentCallback.rowMaps[3]["D"] == "emily@example.example"
        rowContentCallback.rowMaps[3]["E"] == "2/1/1985"
        rowContentCallback.rowMaps[3]["F"] == "9500"

        reader.sheetContentsHandler.columnHeaders["A"]== "Person Id"
        reader.sheetContentsHandler.columnHeaders["B"]== "Name"
        reader.sheetContentsHandler.columnHeaders["C"]== "Height"
        reader.sheetContentsHandler.columnHeaders["D"]== "Email Address"
        reader.sheetContentsHandler.columnHeaders["E"]== "DOB"
        reader.sheetContentsHandler.columnHeaders["F"]== "Salary"

        rowContentCallback.rowMaps[4]["A"] == "10001"
        rowContentCallback.rowMaps[4]["B"] == "Jacob"
        rowContentCallback.rowMaps[4]["C"] == "5,20"
        rowContentCallback.rowMaps[4]["D"] == "jacob@example.example"
        rowContentCallback.rowMaps[4]["E"] == "1/1/1970"
        rowContentCallback.rowMaps[4]["F"] == null

        rowContentCallback.rowMaps[4]
        rowContentCallback.rowMaps[5]
        rowContentCallback.rowMaps[6]
        rowContentCallback.rowMaps[7]
        !rowContentCallback.rowMaps[8]
    }

    void "Read a sample file with 2 times same header column should skip first row and call the handlers with excel column reference as headers "() {

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
        // We consider all columns in header mode

        workSheetHandler.mappingMode = ExcelWorkSheetHandler.ColumnMapping.REFERENCE_AND_HEADER_ROW

        // We skip first row
        workSheetHandler.rowsToSkip = [0]

        // We consider second row  as header
        workSheetHandler.headerRow = 1

        ExcelReader reader = new ExcelReader(file, workSheetHandler)
        reader.sheetCallback = sheetCallback
        reader.sheetSkippedCallback = sheetSkippedCallback
        reader.readCallback = readCallback

        //Ignore sheet 2 & 3
        reader.sheetsToSkip = ["Lot 1 Data", "Lot 2 Data", "Lot 4 Data"]

        when:
        reader.run([:])

        then:
        readCallback.numberOfStarts == 1
        readCallback.numberOfEnds == 1
        sheetCallback.numberOfStarts == 1
        sheetCallback.numberOfEnds == 1
        sheetSkippedCallback.numberOfCalls == 3
        rowContentCallback.numberOfCalls == 3
        rowskippedCallback.numberOfCalls == 1

        rowContentCallback.rowMaps.size() == 3

        Map<String, String> mapFieldValue = rowContentCallback.rowMaps[2]
        mapFieldValue.size()==7

        rowContentCallback.rowMaps[2]["A"] == "30004"
        rowContentCallback.rowMaps[2]["B"] == "Chris"
        rowContentCallback.rowMaps[2]["C"] == "5,50"
        rowContentCallback.rowMaps[2]["D"] == "chris@example.example"
        rowContentCallback.rowMaps[2]["E"] == "8/1/75"
        rowContentCallback.rowMaps[2]["F"] == "8500"
        rowContentCallback.rowMaps[2]["G"] == "10"

        reader.sheetContentsHandler.columnHeaders["A"]== "Person Id"
        reader.sheetContentsHandler.columnHeaders["B"]== "Name"
        reader.sheetContentsHandler.columnHeaders["C"]== "Height"
        reader.sheetContentsHandler.columnHeaders["D"]== "Email Address"
        reader.sheetContentsHandler.columnHeaders["E"]== "DOB"
        reader.sheetContentsHandler.columnHeaders["F"]== "Salary"
        reader.sheetContentsHandler.columnHeaders["G"]== "Salary" // Salary is on column F & G

        rowContentCallback.rowMaps[3]
        rowContentCallback.rowMaps[4]
        !rowContentCallback.rowMaps[5]
    }

    void "Read a sample file should call the handlers with excel column headers"() {

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
        workSheetHandler.rowsToSkip = []

        ExcelReader reader = new ExcelReader(file, workSheetHandler)
        reader.sheetCallback = sheetCallback
        reader.sheetSkippedCallback = sheetSkippedCallback
        reader.readCallback = readCallback

        //Ignore 3 of the 4 sheets
        reader.sheetsToSkip = ["Lot 1 Data", "Lot 2 Data", "Lot 3 Data"]

        when:
        reader.run([:])

        then:
        readCallback.numberOfStarts == 1
        readCallback.numberOfEnds == 1
        sheetCallback.numberOfStarts == 1
        sheetCallback.numberOfEnds == 1
        sheetSkippedCallback.numberOfCalls == 3
        rowContentCallback.numberOfCalls == 2
        rowskippedCallback.numberOfCalls == 0

        rowContentCallback.rowMaps.size() == 2

        Map<String, String> mapFieldValue = rowContentCallback.rowMaps[1]
        mapFieldValue.size()==5
        rowContentCallback.rowMaps[0]["Person Id"] == "20001"
        rowContentCallback.rowMaps[0]["Name"] == "Jacob"
        rowContentCallback.rowMaps[0]["Height"] == null
        rowContentCallback.rowMaps[0]["Email Address"] == "jacob@example.example"
//        rowContentCallback.rowMaps[1]["DOB"] == "1-janv.-1970"
        rowContentCallback.rowMaps[0]["Salary"] == "10500"

        rowContentCallback.rowMaps[1]
        !rowContentCallback.rowMaps[2]
    }

}