package com.softwood.xls

import groovy.json.JsonOutput
import groovy.json.StreamingJsonBuilder
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Header
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow

/**
 * Created by willw on 04/05/2017.
 */

//@Grab('org.apache.poi:poi:3.8')
//@Grab('org.apache.poi:poi-ooxml:3.8')
//@GrabExclude('xml-apis:xml-apis')
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFSheet

import java.util.concurrent.ConcurrentHashMap

def excelFile = new File('Workbook1.xlsx')
excelFile.withInputStream { is ->
    workbook = new XSSFWorkbook(is)
    (0..<workbook.numberOfSheets).each { sheetNumber ->
        XSSFSheet sheet = workbook.getSheetAt(sheetNumber)
        XSSFRow  headerRow = sheet.getRow(0)
        def firstheadercellnum = headerRow.getFirstCellNum()
        def lastheadercellnum = headerRow.getLastCellNum()  //starts at 0 but finishes on colmn past
        Map headers = new ConcurrentHashMap()
        //put each cell with value in headers map
        println "1st cell $firstheadercellnum to last cell $lastheadercellnum"
        for (int hidx in firstheadercellnum..<lastheadercellnum) {
            println "hidx : " + hidx
            def val = headerRow.getCell(hidx)
            println "val = $val at $hidx"
            headers."$hidx" = headerRow.getCell(hidx)
        }
        println "-- done headers : $headers"

        Map outputMap = new ConcurrentHashMap()
        StringWriter writer = new StringWriter()
        StreamingJsonBuilder builder = new StreamingJsonBuilder(writer)
        sheet.rowIterator().each { XSSFRow row ->
            def columnCount = row.physicalNumberOfCells
            def firstcellnum = row.getFirstCellNum()
            def lastcellnum = row.getLastCellNum()
            def nullpad = ""
            (0..<firstcellnum).each  {
                nullpad = nullpad + "null,"
            }
            print nullpad

            if (row.getRowNum() == 0) return //skip header

            Map cellMap = new ConcurrentHashMap()
            row.cellIterator().each { XSSFCell cell ->
                def rowIndex = cell.getRow().getRowNum()
                def colIndex = cell.columnIndex
                def cellType = cell.getCellType()
                def cellValue = cell.toString()
                def headerKey = headers."$colIndex"  //lookup mathced header column
                println "rowIndex : $rowIndex, colIndex : $colIndex, celltype: $cellType, cellValue : $cellValue, headerkey: $headerKey"
                cellMap[headerKey] = cellValue
                //def seperator = (colIndex + 1 < columnCount) ? ",": "\n"
                //print cell.toString() + seperator
            }
            //builder.records {
                //"row$rowIndex" {cellMap.toString()}
            //}
            def op = JsonOutput.toJson(cellMap)
            println JsonOutput.prettyPrint (op)
        }

        //String json = JsonOutput.prettyPrint( writer.toString())
        //println "as json \n" + json
    }
}