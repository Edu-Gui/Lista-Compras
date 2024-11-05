package com.example.lista_compras

import android.content.Intent
import android.os.Bundle
import android.widget.Button
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat

class MainActivity : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        val btnInsertData: Button = findViewById(R.id.btn_insert_data)
        val btnViewExcel: Button = findViewById(R.id.btn_view_excel)

        btnInsertData.setOnClickListener {
            val intent = Intent(this, InsertDataActivity::class.java)
            startActivity(intent)
        }

        btnViewExcel.setOnClickListener {
            val intent = Intent(this, ViewExcelActivity::class.java)
            startActivity(intent)
        }
    }
}
