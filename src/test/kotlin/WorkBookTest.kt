import arrow.core.Tuple6
import arrow.core.getOrHandle
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.jupiter.api.Test
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter
import kotlin.test.assertEquals

class WorkBookTest {

    companion object {
        private val workbook: XSSFWorkbook =
            openResource("workbook.xlsx")
                .map(::openWorkBook)
                .getOrHandle { throw it }

        private val sheet: XSSFSheet = workbook.getSheet("Order")
    }

    @Test
    fun `Extracts table by fixed number of rows`() {
        val table = sheet.toTable(startRow = 4, startColumn = 0, columnCount = 6, rowCount = 4)
        tableOk(table)
    }

    @Test
    fun `Extracts table by row condition`() {
        val digits = "[0-9]+".toRegex()
        val table = sheet.toTable(startRow = 4, startColumn = 0, columnCount = 6) { row ->
            digits.matches(row.cellAsString(column = 0))
        }
        tableOk(table)
    }

    private fun tableOk(table: Table) {
        assertEquals(
            listOf("Item Code", "Item Desc", "Unit Price", "Quantity", "Item Total", "ETA"),
            table.labels
        )
        assertEquals(
            listOf(
                Tuple6(1.0, "Hammer", 2.75, 1.0, 2.75, parseDate("2022-01-04")),
                Tuple6(2.0, "Nails", 0.02, 45.0, 0.9, parseDate("2022-01-05")),
                Tuple6(3.0, "Screwdriver", 1.8, 1.0, 1.8, parseDate("2022-01-06")),
            ),
            table.rows.map { row ->
                println("row: $row")
                Tuple6(
                    row["Item Code"] as Double,
                    row["Item Desc"] as String,
                    row["Unit Price"] as Double,
                    row["Quantity"] as Double,
                    row["Item Total"] as Double,
                    (row["ETA"] as LocalDateTime).toLocalDate()
                )
            }
        )
    }

    private fun parseDate(s: String) = LocalDate.parse(s, DateTimeFormatter.ISO_LOCAL_DATE)
}
