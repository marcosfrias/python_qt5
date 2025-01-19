import ctypes
import os
import sys
import res_rc
import res_btn_rc
import pandas as pd
from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QFileDialog, QTableWidgetItem, QPushButton
from drop_button import DropButton
import qdarktheme
from table_window import TableWindow


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('temp.ui', self)

        # Reemplazo de Widget clase personalizada
        self.pushButton_drop = DropButton(parent=self)
        self.widget_sub.layout().addWidget(self.pushButton_drop)

        # Configurar el tema oscuro
        app.setStyleSheet(qdarktheme.load_stylesheet("light"))

        # Conecta el botón de subir archivo
        self.pushButton_drop.clicked.connect(self.open_file)

        # Conecta el search_btn con la función de búsqueda
        self.search_button = self.findChild(QtWidgets.QPushButton, "search_btn")
        self.search_button.clicked.connect(self.open_table_window)

        # Conecta el lineEdit con la función de búsqueda en tiempo real
        self.lineEdit.textChanged.connect(self.search_value)

        # Configurar ícono en la barra de tareas (Windows)
        myappid = 'mi.aplicacion.id.version.1.0'  # Identificador único
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    def open_file(self):
        options = QtWidgets.QFileDialog.Options()
        file_name, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Abrir archivo Excel", "",
            "Archivos Excel (*.xls *.xlsx);;Todos los archivos (*)", options=options
        )
        if file_name:
            self.load_excel(file_name)

    def load_excel(self, file_name):
        try:
            print(f"Intentando leer el archivo Excel: {file_name}")  # Depuración
            self.df = pd.read_excel(file_name)
            print(f"Archivo Excel cargado: {file_name}")  # Depuración
            print(f"Estructura del DataFrame:\n{self.df.head()}")  # Depuración
            print("Columnas del DataFrame:", self.df.columns)
            print("Actualizando la tabla...")  # Depuración
            self.tableWidget.setRowCount(self.df.shape[0])
            self.tableWidget.setColumnCount(self.df.shape[1])
            self.tableWidget.setHorizontalHeaderLabels(self.df.columns)

            for i in range(self.df.shape[0]):
                for j in range(self.df.shape[1]):
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(self.df.iat[i, j])))
                    print(f"Celda actualizada: ({i}, {j}) - {self.df.iat[i, j]}")  # Depuración

            file_name_only = os.path.basename(file_name)
            self.label_up.setText(f"{file_name_only}")

            # Oculta el DropButton después de cargar el archivo
            self.pushButton_drop.setVisible(False)

            # Establece un ancho específico para las columnas del tableWidget
            for column in range(self.tableWidget.columnCount()):
                self.tableWidget.setColumnWidth(column, 200)
        except Exception as e:
            self.label_up.setText(f"Error al cargar el archivo: {e}")
            print(f"Error: {e}")  # Depuración

    def search_value(self):
        if hasattr(self, 'df'):  # Verifica si self.df está definido
            value = self.lineEdit.text()
            print(f"Valor ingresado en lineEdit: {value}")  # Depuración
            if 'EMPRESAS' in self.df.columns:
                filtered_df = self.df[self.df['EMPRESAS'].astype(str).str.contains(value, na=False, case=False)]
                print(f"DataFrame filtrado:\n{filtered_df}")  # Depuración
                self.update_table(filtered_df)
            else:
                print("La columna 'EMPRESAS' no existe en el DataFrame.")
                self.label_up.setText("Error: La columna 'EMPRESAS' no existe.")
        else:
            print("El DataFrame no está cargado. Carga un archivo Excel primero.")
            self.label_up.setText("Error: El DataFrame no está cargado. Carga un archivo Excel primero.")

    def update_table(self, df):
        self.tableWidget.setRowCount(df.shape[0])
        self.tableWidget.setColumnCount(df.shape[1])
        self.tableWidget.setHorizontalHeaderLabels(df.columns)

        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))
                print(f"Celda actualizada: ({i}, {j}) - {df.iat[i, j]}")  # Depuración

    def open_table_window(self):
        self.table_window = TableWindow(self.tableWidget)
        self.table_window.show()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
