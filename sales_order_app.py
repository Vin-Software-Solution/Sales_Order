import sys
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QTableWidget, QTableWidgetItem, QTabWidget, 
                             QLabel, QLineEdit, QDateEdit, QFormLayout, QGroupBox, 
                             QMessageBox, QFileDialog, QHeaderView)
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QFont
import sqlite3
import os

class SalesOrderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Aplikasi Manajemen Sales dan Repeat Order")
        self.setGeometry(100, 100, 1000, 700)
        
        # Inisialisasi database
        self.init_db()
        
        # Setup UI
        self.setup_ui()
        
        # Load data
        self.load_sales_data()
        self.load_order_data()
    
    def init_db(self):
        """Initialize database and tables"""
        self.conn = sqlite3.connect('sales_orders.db')
        self.cursor = self.conn.cursor()
        
        # Create sales table if not exists
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                email TEXT,
                phone TEXT,
                join_date TEXT,
                territory TEXT
            )
        ''')
        
        # Create orders table if not exists
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sales_id INTEGER,
                customer_name TEXT NOT NULL,
                product TEXT,
                quantity INTEGER,
                order_date TEXT,
                amount REAL,
                status TEXT,
                FOREIGN KEY (sales_id) REFERENCES sales (id)
            )
        ''')
        
        self.conn.commit()
    
    def setup_ui(self):
        """Setup the user interface"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Create tabs
        tabs = QTabWidget()
        layout.addWidget(tabs)
        
        # Sales tab
        sales_tab = QWidget()
        sales_layout = QVBoxLayout(sales_tab)
        
        # Sales form
        sales_form_group = QGroupBox("Form Data Sales")
        sales_form_layout = QFormLayout()
        
        self.sales_name = QLineEdit()
        self.sales_email = QLineEdit()
        self.sales_phone = QLineEdit()
        self.sales_join_date = QDateEdit()
        self.sales_join_date.setDate(QDate.currentDate())
        self.sales_territory = QLineEdit()
        
        sales_form_layout.addRow(QLabel("Nama:"), self.sales_name)
        sales_form_layout.addRow(QLabel("Email:"), self.sales_email)
        sales_form_layout.addRow(QLabel("Telepon:"), self.sales_phone)
        sales_form_layout.addRow(QLabel("Tanggal Bergabung:"), self.sales_join_date)
        sales_form_layout.addRow(QLabel("Wilayah:"), self.sales_territory)
        
        sales_form_group.setLayout(sales_form_layout)
        sales_layout.addWidget(sales_form_group)
        
        # Sales buttons
        sales_buttons_layout = QHBoxLayout()
        self.add_sales_btn = QPushButton("Tambah Sales")
        self.add_sales_btn.clicked.connect(self.add_sales)
        self.import_sales_btn = QPushButton("Impor dari Excel/CSV")
        self.import_sales_btn.clicked.connect(self.import_sales_data)
        
        sales_buttons_layout.addWidget(self.add_sales_btn)
        sales_buttons_layout.addWidget(self.import_sales_btn)
        sales_layout.addLayout(sales_buttons_layout)
        
        # Sales table
        self.sales_table = QTableWidget()
        self.sales_table.setColumnCount(6)
        self.sales_table.setHorizontalHeaderLabels(["ID", "Nama", "Email", "Telepon", "Tanggal Bergabung", "Wilayah"])
        self.sales_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        sales_layout.addWidget(self.sales_table)
        
        # Orders tab
        orders_tab = QWidget()
        orders_layout = QVBoxLayout(orders_tab)
        
        # Orders form
        orders_form_group = QGroupBox("Form Repeat Order")
        orders_form_layout = QFormLayout()
        
        self.order_sales = QLineEdit()
        self.order_customer = QLineEdit()
        self.order_product = QLineEdit()
        self.order_quantity = QLineEdit()
        self.order_date = QDateEdit()
        self.order_date.setDate(QDate.currentDate())
        self.order_amount = QLineEdit()
        self.order_status = QLineEdit()
        
        orders_form_layout.addRow(QLabel("ID Sales:"), self.order_sales)
        orders_form_layout.addRow(QLabel("Nama Pelanggan:"), self.order_customer)
        orders_form_layout.addRow(QLabel("Produk:"), self.order_product)
        orders_form_layout.addRow(QLabel("Jumlah:"), self.order_quantity)
        orders_form_layout.addRow(QLabel("Tanggal Order:"), self.order_date)
        orders_form_layout.addRow(QLabel("Jumlah (Rp):"), self.order_amount)
        orders_form_layout.addRow(QLabel("Status:"), self.order_status)
        
        orders_form_group.setLayout(orders_form_layout)
        orders_layout.addWidget(orders_form_group)
        
        # Orders buttons
        orders_buttons_layout = QHBoxLayout()
        self.add_order_btn = QPushButton("Tambah Order")
        self.add_order_btn.clicked.connect(self.add_order)
        self.import_orders_btn = QPushButton("Impor dari Excel/CSV")
        self.import_orders_btn.clicked.connect(self.import_order_data)
        
        orders_buttons_layout.addWidget(self.add_order_btn)
        orders_buttons_layout.addWidget(self.import_orders_btn)
        orders_layout.addLayout(orders_buttons_layout)
        
        # Orders table
        self.orders_table = QTableWidget()
        self.orders_table.setColumnCount(8)
        self.orders_table.setHorizontalHeaderLabels(["ID", "ID Sales", "Pelanggan", "Produk", "Jumlah", "Tanggal", "Jumlah (Rp)", "Status"])
        self.orders_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        orders_layout.addWidget(self.orders_table)
        
        # Add tabs to the tab widget
        tabs.addTab(sales_tab, "Data Sales")
        tabs.addTab(orders_tab, "Repeat Order")
    
    def add_sales(self):
        """Add a new sales person to the database"""
        name = self.sales_name.text()
        email = self.sales_email.text()
        phone = self.sales_phone.text()
        join_date = self.sales_join_date.date().toString("yyyy-MM-dd")
        territory = self.sales_territory.text()
        
        if not name:
            QMessageBox.warning(self, "Peringatan", "Nama sales harus diisi!")
            return
        
        try:
            self.cursor.execute(
                "INSERT INTO sales (name, email, phone, join_date, territory) VALUES (?, ?, ?, ?, ?)",
                (name, email, phone, join_date, territory)
            )
            self.conn.commit()
            
            # Clear form
            self.sales_name.clear()
            self.sales_email.clear()
            self.sales_phone.clear()
            self.sales_territory.clear()
            
            # Reload data
            self.load_sales_data()
            
            QMessageBox.information(self, "Sukses", "Data sales berhasil ditambahkan!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Terjadi kesalahan: {str(e)}")
    
    def add_order(self):
        """Add a new order to the database"""
        try:
            sales_id = int(self.order_sales.text())
        except ValueError:
            QMessageBox.warning(self, "Peringatan", "ID Sales harus angka!")
            return
            
        customer = self.order_customer.text()
        product = self.order_product.text()
        
        try:
            quantity = int(self.order_quantity.text())
        except ValueError:
            QMessageBox.warning(self, "Peringatan", "Jumlah harus angka!")
            return
            
        order_date = self.order_date.date().toString("yyyy-MM-dd")
        
        try:
            amount = float(self.order_amount.text())
        except ValueError:
            QMessageBox.warning(self, "Peringatan", "Jumlah (Rp) harus angka!")
            return
            
        status = self.order_status.text()
        
        if not customer:
            QMessageBox.warning(self, "Peringatan", "Nama pelanggan harus diisi!")
            return
        
        try:
            self.cursor.execute(
                "INSERT INTO orders (sales_id, customer_name, product, quantity, order_date, amount, status) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (sales_id, customer, product, quantity, order_date, amount, status)
            )
            self.conn.commit()
            
            # Clear form
            self.order_sales.clear()
            self.order_customer.clear()
            self.order_product.clear()
            self.order_quantity.clear()
            self.order_amount.clear()
            self.order_status.clear()
            
            # Reload data
            self.load_order_data()
            
            QMessageBox.information(self, "Sukses", "Data order berhasil ditambahkan!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Terjadi kesalahan: {str(e)}")
    
    def load_sales_data(self):
        """Load sales data from database to table"""
        self.cursor.execute("SELECT * FROM sales")
        sales_data = self.cursor.fetchall()
        
        self.sales_table.setRowCount(len(sales_data))
        for row_idx, row_data in enumerate(sales_data):
            for col_idx, col_data in enumerate(row_data):
                self.sales_table.setItem(row_idx, col_idx, QTableWidgetItem(str(col_data)))
    
    def load_order_data(self):
        """Load order data from database to table"""
        self.cursor.execute("SELECT * FROM orders")
        order_data = self.cursor.fetchall()
        
        self.orders_table.setRowCount(len(order_data))
        for row_idx, row_data in enumerate(order_data):
            for col_idx, col_data in enumerate(row_data):
                self.orders_table.setItem(row_idx, col_idx, QTableWidgetItem(str(col_data)))
    
    def import_sales_data(self):
        """Import sales data from Excel or CSV file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Pilih File Excel atau CSV", "", "Excel/CSV Files (*.xlsx *.xls *.csv)"
        )
        
        if not file_path:
            return
        
        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            
            # Process each row
            for _, row in df.iterrows():
                # Handle missing values
                name = row.get('name', '') or row.get('nama', '') or ''
                email = row.get('email', '') or ''
                phone = row.get('phone', '') or row.get('telepon', '') or ''
                join_date = row.get('join_date', '') or row.get('tanggal_bergabung', '') or QDate.currentDate().toString("yyyy-MM-dd")
                territory = row.get('territory', '') or row.get('wilayah', '') or ''
                
                if name:  # Only insert if name exists
                    self.cursor.execute(
                        "INSERT INTO sales (name, email, phone, join_date, territory) VALUES (?, ?, ?, ?, ?)",
                        (name, email, phone, join_date, territory)
                    )
            
            self.conn.commit()
            self.load_sales_data()
            QMessageBox.information(self, "Sukses", f"Data sales berhasil diimpor dari {os.path.basename(file_path)}!")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat mengimpor: {str(e)}")
    
    def import_order_data(self):
        """Import order data from Excel or CSV file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Pilih File Excel atau CSV", "", "Excel/CSV Files (*.xlsx *.xls *.csv)"
        )
        
        if not file_path:
            return
        
        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            
            # Process each row
            for _, row in df.iterrows():
                # Handle missing values
                sales_id = row.get('sales_id', 0) or 0
                customer = row.get('customer_name', '') or row.get('pelanggan', '') or ''
                product = row.get('product', '') or row.get('produk', '') or ''
                quantity = row.get('quantity', 0) or row.get('jumlah', 0) or 0
                order_date = row.get('order_date', '') or row.get('tanggal_order', '') or QDate.currentDate().toString("yyyy-MM-dd")
                amount = row.get('amount', 0) or row.get('jumlah_rp', 0) or 0
                status = row.get('status', '') or ''
                
                if customer:  # Only insert if customer exists
                    self.cursor.execute(
                        "INSERT INTO orders (sales_id, customer_name, product, quantity, order_date, amount, status) VALUES (?, ?, ?, ?, ?, ?, ?)",
                        (sales_id, customer, product, quantity, order_date, amount, status)
                    )
            
            self.conn.commit()
            self.load_order_data()
            QMessageBox.information(self, "Sukses", f"Data order berhasil diimpor dari {os.path.basename(file_path)}!")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat mengimpor: {str(e)}")
    
    def closeEvent(self, event):
        """Close database connection when application closes"""
        self.conn.close()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SalesOrderApp()
    window.show()
    sys.exit(app.exec_())