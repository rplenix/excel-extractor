import arrow.core.Tuple6
import arrow.core.getOrHandle
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.jupiter.api.Test
import java.text.SimpleDateFormat
import java.util.Date
import kotlin.test.assertEquals

class WorkBookTest {

    private val workbook: XSSFWorkbook =
        openResource("workbook.xlsx")
            .map(::openWorkBook)
            .getOrHandle { throw it }

    private val sheet: XSSFSheet = workbook.getSheet("Order")

    @Test
    fun `Extracts table`() {
        val table = sheet.toTable(4, 0, 6, 4)
        assertEquals(
            listOf("Item Code", "Item Desc", "Unit Price", "Quantity", "Item Total", "ETA"),
            table.labels
        )
        val dateFormat = SimpleDateFormat("yyyy-MM-dd")
        assertEquals(
            listOf(
                Tuple6(1.0, "Hammer", 2.75, 1.0, 2.75, dateFormat.parse("2022-01-04")),
                Tuple6(2.0, "Nails", 0.02, 45.0, 0.9, dateFormat.parse("2022-01-05")),
                Tuple6(3.0, "Screwdriver", 1.8, 1.0, 1.8, dateFormat.parse("2022-01-06")),
            ),
            table.rows.map { row ->
                println("row: $row")
                Tuple6(
                    row["Item Code"] as Double,
                    row["Item Desc"] as String,
                    row["Unit Price"] as Double,
                    row["Quantity"] as Double,
                    row["Item Total"] as Double,
                    row["ETA"] as Date
                )
            }
        )
        println(table)
    }
}
