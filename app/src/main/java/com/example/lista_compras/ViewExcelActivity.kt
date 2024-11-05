package com.example.lista_compras
import android.os.Bundle
import android.widget.TableLayout
import android.widget.TableRow
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileInputStream
import java.io.IOException

class ViewExcelActivity : AppCompatActivity() {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_view_excel)

        loadExcelData()
    }

    private fun loadExcelData() {
        val file = File(getExternalFilesDir(null), "planilha.xlsx")
        if (!file.exists()) {
            Toast.makeText(this, "Nenhuma planilha encontrada.", Toast.LENGTH_SHORT).show()
            return
        }

        val tableLayout: TableLayout = findViewById(R.id.tl_excel_data)

        try {
            FileInputStream(file).use { fis ->
                val workbook = WorkbookFactory.create(fis)
                val sheet: Sheet? = workbook.getSheet("Clientes")

                if (sheet == null) {
                    Toast.makeText(this, "A aba 'Clientes' não foi encontrada na planilha.", Toast.LENGTH_SHORT).show()
                    workbook.close()
                    return
                }

                tableLayout.removeAllViews()

                for (i in 0 until sheet.physicalNumberOfRows) {
                    val row = sheet.getRow(i)
                    val tableRow = TableRow(this)

                    for (j in 0 until row.physicalNumberOfCells) {
                        val cell = row.getCell(j)
                        val textView = TextView(this)
                        textView.text = cell.toString()
                        textView.setPadding(16, 16, 16, 16)
                        tableRow.addView(textView)
                    }

                    tableLayout.addView(tableRow)
                }

                workbook.close()
            }
        } catch (e: IOException) {
            Toast.makeText(this, "Erro ao carregar a planilha: falha de I/O.", Toast.LENGTH_SHORT).show()
            e.printStackTrace()
        } catch (e: Exception) {
            Toast.makeText(this, "Erro ao carregar a planilha.", Toast.LENGTH_SHORT).show()
            e.printStackTrace()
        }
    }

}


