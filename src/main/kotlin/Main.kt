
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream

fun main() {
    try {
        val file = File("AB.xlsx")
        val fis = FileInputStream(file)
        val workbook = XSSFWorkbook(fis)
        val sheet: XSSFSheet = workbook.getSheetAt(0)

        // Iterate through each row
        for (row in sheet) {
            val cellIterator = row.cellIterator()
            var concatenatedValue = "" // Initialize an empty string

            while (cellIterator.hasNext()) {
                val cell: Cell = cellIterator.next()

                when (cell.cellType) {
                    CellType.STRING -> {
                        val cellValue = cell.stringCellValue;
                    }
                    CellType.NUMERIC -> {
                        val numericValue = cell.numericCellValue.toInt()
                        val charValue = ('A' + numericValue - 1).toChar() // Convert numeric value to corresponding character
                        concatenatedValue += "$numericValue$charValue " // Append the combined value
                    }
                    else -> {
                        // Handle other cell types if needed
                    }
                }
            }

            // Set the concatenated value in the third column (index 2)
            val thirdCell = row.createCell(2)
            thirdCell.setCellValue(concatenatedValue)
        }

        // Save the modified workbook
        val outputStream = FileOutputStream("Modified_AB.xlsx")
        workbook.write(outputStream)
        workbook.close()
        fis.close()

        println("Concatenated values saved to Modified_AB.xlsx")
    } catch (e: Exception) {
        e.printStackTrace()
    }
}
