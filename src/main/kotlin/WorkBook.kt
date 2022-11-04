import org.apache.poi.ss.usermodel.CellType.BOOLEAN
import org.apache.poi.ss.usermodel.CellType.FORMULA
import org.apache.poi.ss.usermodel.CellType.NUMERIC
import org.apache.poi.ss.usermodel.CellType.STRING
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.InputStream

fun openWorkBook(inputStream: InputStream) = XSSFWorkbook(inputStream)

data class Row(val columnValues: Map<String, Any?>) {
    operator fun get(name: String): Any? = columnValues[name]
}

data class Table(val labels: List<String>, val rows: List<Row>) {
    val size: Int = rows.size
    operator fun get(i: Int) = rows[i]
}

fun XSSFSheet.toTable(startRow: Int, startColumn: Int, columnCount: Int, rowCount: Int): Table {
    val labels =
        (startColumn until startColumn + columnCount).map { columnIndex -> cellAsString(startRow, columnIndex) }
    val rows =
        (startRow + 1 until startRow + rowCount)
            .map { rowIndex ->
                labels.indices.associate { columnIndex ->
                    Pair(labels[columnIndex], cellAsAny(rowIndex, startColumn + columnIndex))
                }
            }
            .map(::Row)
    return Table(labels, rows)
}

fun XSSFSheet.cellAsAny(row: Int, column: Int): Any? =
    getRow(row)?.getCell(column)?.let { cell ->
        when (cell.cellType) {
            STRING -> cell.stringCellValue
            FORMULA -> {
                val cellValue = workbook.creationHelper.createFormulaEvaluator().evaluate(cell)
                when (cell.cachedFormulaResultType) {
                    NUMERIC -> {
                        if (DateUtil.isCellDateFormatted(cell)) {
                            cell.dateCellValue
                        } else {
                            cellValue.numberValue
                        }
                    }

                    BOOLEAN -> cellValue.booleanValue
                    STRING -> cellValue.stringValue
                    else -> cellValue.toString()
                }
            }

            NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    cell.dateCellValue
                } else {
                    cell.numericCellValue
                }
            }
            BOOLEAN -> cell.booleanCellValue
            else -> null
        }
    }

fun XSSFSheet.cellAsString(row: Int, column: Int): String =
    getRow(row)
        ?.getCell(column)
        ?.let { cell ->
            when (cell.cellType) {
                STRING -> cell.stringCellValue
                FORMULA -> DataFormatter().formatCellValue(cell, workbook.creationHelper.createFormulaEvaluator())
                else -> DataFormatter().formatCellValue(cell)
            }
        }
        ?: ""
