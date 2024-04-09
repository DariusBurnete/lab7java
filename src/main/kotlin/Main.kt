// Import statements
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream

fun main() {

            // Try block to check for exceptions
            try {

                val file = File("AB.xlsx")
                FileInputStream(file)

                // Create Workbook instance holding reference to
                // .xlsx file
                val workbook = XSSFWorkbook(file)

                // Get first/desired sheet from the workbook
                val sheet: XSSFSheet = workbook.getSheetAt(0)


                // Iterate through each rows one by one
                val rowIterator: Iterator<Row> = sheet.iterator()

                var sum = 0.0 // Initialize the sum variable

                while (rowIterator.hasNext()) {
                    val row: Row = rowIterator.next()
                    val cellIterator: Iterator<Cell> = row.cellIterator()

                    while (cellIterator.hasNext()) {
                        val cell: Cell = cellIterator.next()

                        when (cell.cellType) {
                            CellType.STRING -> {
                                val cellValue = cell.stringCellValue
                                // Handle string cell value
                                print("$cellValue ")
                            }
                            CellType.NUMERIC -> {
                                val numericValue = cell.numericCellValue
                                // Handle numeric cell value
                                print("$numericValue ")
                                sum += numericValue // Add to the sum
                            }
                            else -> {
                                // Handle other cell types (e.g., BOOLEAN, FORMULA, BLANK, etc.)
                                print("Unknown cell type ")
                            }
                        }
                    }

                    println("") // Print a newline after each row
                }

                println("Sum of numeric values: $sum") // Print the total sum

                file.inputStream().use {
                    // Close the input stream (file) automatically
                }
            }

            // Catch block to handle exceptions
            catch (e: Exception) {
                e.printStackTrace()
            }
        }
