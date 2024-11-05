package com.example.lista_compras

import android.os.Bundle
import android.text.InputFilter
import android.widget.Button
import android.widget.EditText
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException
import java.text.SimpleDateFormat
import java.util.Calendar

class InsertDataActivity : AppCompatActivity() {

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_insert_data)

        val etNomeCliente: EditText = findViewById(R.id.et_client_name)
        val etTelefone: EditText = findViewById(R.id.et_phone)
        val etProduto: EditText = findViewById(R.id.et_product)
        val etValor: EditText = findViewById(R.id.et_value)
        val etFormaPagamento: EditText = findViewById(R.id.et_payment_method)
        val etDataPagamento: EditText = findViewById(R.id.et_payment_date)
        val etProximaData: EditText = findViewById(R.id.et_next_payment_date)
        val btnSave: Button = findViewById(R.id.btn_save)

        val onlyLettersFilter = InputFilter { source, start, end, dest, dstart, dend ->
            if (source.matches(Regex("[a-zA-Z ]+"))) null else ""
        }
        val onlyNumbersFilter = InputFilter { source, start, end, dest, dstart, dend ->
            if (source.matches(Regex("[0-9]+"))) null else ""
        }
        val decimalNumbersFilter = InputFilter { source, start, end, dest, dstart, dend ->
            if (source.matches(Regex("[0-9.]+"))) null else ""
        }

        etNomeCliente.filters = arrayOf(onlyLettersFilter)
        etProduto.filters = arrayOf(onlyLettersFilter)
        etTelefone.filters = arrayOf(onlyNumbersFilter)
        etValor.filters = arrayOf(decimalNumbersFilter)

        val dateRegex = Regex("""\d{2}/\d{2}/\d{4}""")

        btnSave.setOnClickListener {
            val nomeCliente = etNomeCliente.text.toString()
            val telefone = etTelefone.text.toString()
            val produto = etProduto.text.toString()
            val valor = etValor.text.toString()
            val formaPagamento = etFormaPagamento.text.toString()
            val dataPagamento = etDataPagamento.text.toString()
            val proximaData = etProximaData.text.toString()

            var isValid = true

            if (nomeCliente.isEmpty()) {
                etNomeCliente.error = "Nome é obrigatório"
                isValid = false
            }
            if (telefone.isEmpty()) {
                etTelefone.error = "Telefone é obrigatório"
                isValid = false
            }
            if (produto.isEmpty()) {
                etProduto.error = "Produto é obrigatório"
                isValid = false
            }
            if (valor.isEmpty()) {
                etValor.error = "Valor é obrigatório"
                isValid = false
            }
            if (!dateRegex.matches(dataPagamento)) {
                etDataPagamento.error = "Data inválida. Use o formato dd/MM/yyyy"
                isValid = false
            }
            if (!dateRegex.matches(proximaData)) {
                etProximaData.error = "Data inválida. Use o formato dd/MM/yyyy"
                isValid = false
            }

            if (isValid) {
                salvarDadosNoExcel(
                    nomeCliente, telefone, produto, valor,
                    formaPagamento, dataPagamento, proximaData
                )
            } else {
                Toast.makeText(this, "Corrija os erros antes de salvar", Toast.LENGTH_SHORT).show()
            }
        }
    }

    private fun salvarDadosNoExcel(
        nomeCliente: String, telefone: String, produto: String, valor: String,
        formaPagamento: String, dataPagamento: String, proximaData: String
    ) {
        val file = File(getExternalFilesDir(null), "planilha.xlsx")
        var workbook: Workbook? = null
        var sheet: Sheet? = null

        try {
            if (file.exists() && file.length() > 0) {
                FileInputStream(file).use { fis ->
                    workbook = WorkbookFactory.create(fis)
                }
            } else {
                workbook = XSSFWorkbook()
            }

            sheet = workbook?.getSheet("Clientes") ?: workbook?.createSheet("Clientes")

            val row: Row = sheet?.createRow(sheet.physicalNumberOfRows) ?: return
            row.createCell(0).setCellValue(nomeCliente)
            row.createCell(1).setCellValue(telefone)
            row.createCell(2).setCellValue(produto)
            row.createCell(3).setCellValue(valor)
            row.createCell(4).setCellValue(formaPagamento)
            row.createCell(5).setCellValue(dataPagamento)
            row.createCell(6).setCellValue(proximaData)

            val currentTime = getCurrentDateTime()
            row.createCell(7).setCellValue(currentTime)


            FileOutputStream(file).use { fos ->
                workbook?.write(fos)
                fos.flush()
            }

            Toast.makeText(this, "Dados salvos com sucesso", Toast.LENGTH_SHORT).show()

            finish()

        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this, "Erro ao salvar a planilha.", Toast.LENGTH_SHORT).show()
        } finally {
            try {
                workbook?.close()
            } catch (e: IOException) {
                e.printStackTrace()
            }
        }
    }

    private fun getCurrentDateTime(): String {
        val calendar = Calendar.getInstance()
        val dateFormat = SimpleDateFormat("dd/MM/yyyy HH:mm:ss")
        return dateFormat.format(calendar.time)
    }
}



