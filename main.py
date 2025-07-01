import sys
from PyQt5.uic import loadUi
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QDialog, QApplication

class Information:
    """Clase para almacenar información"""
    def __init__(self):
        self.data = {}  # Use a dictionary to store key-value pairs
    def set_value(self, key, value):
        """Set a value for a given key"""
        
        self.data[key] = value
        print(f"-> set \"{self.data}\"")
    def get_value(self, key):
        """Get a value for a given key"""
        print(f"-> get \"{self.data}\"")
        return self.data.get(key, "")  # Return an empty string if the key does not exist

class WindowManager:
    """Clase para administrar la transicion entre ventanas"""
    def __init__(self, info):
        self.info = info
        self.current_window = None

    def show_window(self, window_class):
        """Mostrar una nueva ventana y ocultar la actual"""
        if self.current_window is not None:
            self.current_window.hide()  # Hide the current window
        self.current_window = window_class(self.info)  # Create a new window instance
        self.current_window.show()  # Show the new window

    def setup_comboBox(self, comboBox, elementoslista, key, defaultindex):
        """Configurar comboBox con los elementos dados y almacenar el valor seleccionado"""
        # Rellenar valores de comboBox
        comboBox.addItems(elementoslista)
        
        # Recuperar valores almacenados
        value = self.info.get_value(key)
        index = elementoslista.index(value) if value in elementoslista else defaultindex
        
        # Obtener el index actual y almacenar el default value
        comboBox.setCurrentIndex(index)
        self.info.set_value(key, elementoslista[index] if value == "" else value)
        
        # Conecte la señal currentTextChanged para actualizar la clase Information
        comboBox.currentTextChanged.connect(lambda x: self.info.set_value(key, x))   
    
    def setup_lineEdit(self, lineEdit, key):
        """Setup a lineEdit and store its value."""
        default_value = self.info.get_value(key)  # Recuperar el valor almacenado
        lineEdit.setText(default_value)  # Establezca el texto actual de QLineEdit en el valor almacenado
        
        # Conecte la señal textChanged para almacenar el valor de QLineEdit
        lineEdit.textChanged.connect(lambda x: self.info.set_value(key, x))


class Ventana1(QDialog):
    def __init__(self, info):
        super(Ventana1, self).__init__()
        loadUi("ventana1.ui", self)
        self.info = info  # Store the information instance
        self.next_button.clicked.connect(lambda: self.go_to_ventana(Ventana2))  # Connect the button to the method

        # Configuracion de los QComboBox
        window_manager.setup_comboBox(
            comboBox=self.comboBox_2,
            elementoslista=["Muy ligero: 12.7 mm/kV", "Ligero: 16 mm/kV", "Medio: 20 mm/kV", "Pesado: 25 mm/kV", "Muy pesado: 31 mm/kV"],
            key="NIV_DE_CONTAM",
            defaultindex=4  # Default to "Muy pesado"
        )
        # Configuracion de los QLineEdit
        window_manager.setup_lineEdit(
            lineEdit=self.lineEdit,
            key="V_NOM"  # Setup lineEdit with key
        )
        window_manager.setup_lineEdit(
            lineEdit=self.lineEdit_2,
            key="V_MAX"  # Setup lineEdit with key
        )        
        window_manager.setup_lineEdit(
            lineEdit=self.lineEdit_3,
            key="HZ"  # Setup lineEdit with key
        )
        window_manager.setup_lineEdit(
            lineEdit=self.lineEdit_4,
            key="ICC"  # Setup lineEdit with key
        )

    def go_to_ventana(self,ventana):
        window_manager.show_window(ventana)  # Use the window manager to show ventana

class Ventana2(QDialog):
    def __init__(self, info):
        super(Ventana2, self).__init__()
        loadUi("ventana2.ui", self)
        self.info = info  # Store the information instance
        self.prev_button.clicked.connect(lambda: self.go_to_ventana(Ventana1)) # Connect the back button to the method
        self.next_button.clicked.connect(lambda: self.go_to_ventana(Ventana3))  # Connect the next button to a new method
        self.pushButton_linebay.clicked.connect(lambda: self.go_to_ventana(Ventana2_1))

    def go_to_ventana(self,ventana):
        window_manager.show_window(ventana)  # Use the window manager to show ventana

class Ventana2_1(QDialog):
    def __init__(self, info):
        super(Ventana2_1, self).__init__()
        loadUi("ventana2_1.ui", self)
        self.info = info  # Store the information instance
        self.prev_button.clicked.connect(lambda: self.go_to_ventana(Ventana2)) # Connect the back button to the method
        self.pushButton_TT.clicked.connect(lambda: self.go_to_ventana(Ventana2_1_1))
        self.pushButton_TC.clicked.connect(lambda: self.go_to_ventana(Ventana2_1_2))
        self.pushButton_SL.clicked.connect(lambda: self.go_to_ventana(Ventana2_1_3))
        self.pushButton_SB.clicked.connect(lambda: self.go_to_ventana(Ventana2_1_4))
        self.pushButton_IN.clicked.connect(lambda: self.go_to_ventana(Ventana2_1_5))

    def go_to_ventana(self,ventana):
        window_manager.show_window(ventana)  # Use the window manager to show ventana

class Ventana2_1_1(QDialog):
    def __init__(self, info):
        super(Ventana2_1_1, self).__init__()
        loadUi("ventana2_1_1.ui", self)
        self.info = info  # Store the information instance
        self.prev_button.clicked.connect(lambda: self.go_to_ventana(Ventana2_1))

    def go_to_ventana(self,ventana):
        window_manager.show_window(ventana)  # Use the window manager to show ventana

class Ventana2_1_2(QDialog):
    def __init__(self, info):
        super(Ventana2_1_2, self).__init__()
        loadUi("ventana2_1_2.ui", self)
        self.info = info  # Store the information instance
        self.prev_button.clicked.connect(lambda: self.go_to_ventana(Ventana2_1))

    def go_to_ventana(self,ventana):
        window_manager.show_window(ventana)  # Use the window manager to show ventana

class Ventana2_1_3(QDialog):
    def __init__(self, info):
        super(Ventana2_1_3, self).__init__()
        loadUi("ventana2_1_3.ui", self)
        self.info = info  # Store the information instance
        self.prev_button.clicked.connect(lambda: self.go_to_ventana(Ventana2_1))

    def go_to_ventana(self,ventana):
        window_manager.show_window(ventana)  # Use the window manager to show ventana

class Ventana2_1_4(QDialog):
    def __init__(self, info):
        super(Ventana2_1_4, self).__init__()
        loadUi("ventana2_1_4.ui", self)
        self.info = info  # Store the information instance
        self.prev_button.clicked.connect(lambda: self.go_to_ventana(Ventana2_1))

    def go_to_ventana(self,ventana):
        window_manager.show_window(ventana)  # Use the window manager to show ventana

class Ventana2_1_5(QDialog):
    def __init__(self, info):
        super(Ventana2_1_5, self).__init__()
        loadUi("ventana2_1_5.ui", self)
        self.info = info  # Store the information instance
        self.prev_button.clicked.connect(lambda: self.go_to_ventana(Ventana2_1))

    def go_to_ventana(self,ventana):
        window_manager.show_window(ventana)  # Use the window manager to show ventana
class Ventana3(QDialog):
    def __init__(self, info):
        super(Ventana3, self).__init__()
        loadUi("ventana3.ui", self)
        self.info = info  # Store the information instance
        self.prev_button.clicked.connect(lambda: self.go_to_ventana(Ventana2))  # Connect the back button to the method
        self.next_button.clicked.connect(lambda: self.go_to_ventana(Ventana4))  # Connect the next button to a new method
        self.pushButton3.clicked.connect(self.guardar)  # Connect the save button to the method
        self.lineEdit.setText(self.info.get_data())  # Set the text in the line edit

    def go_to_ventana(self,ventana):
        window_manager.show_window(ventana)  # Use the window manager to show ventana

    def guardar(self):
        # Save the data from the line edit to the information instance
        self.info.set_data(self.lineEdit.text())
        print(f"Data saved: {self.info.get_data()}")  # Optional: Print the saved data

class Ventana4(QDialog):
    def __init__(self, info):
        super(Ventana4, self).__init__()
        loadUi("ventana4.ui", self)
        self.info = info  # Store the information instance
        self.prev_button.clicked.connect(lambda: self.go_to_ventana(Ventana3))  # Connect the back button to the method
        self.pushButton3.clicked.connect(self.guardar)  # Connect the save button to the method
        self.lineEdit2.setText(self.info.get_data())  # Set the text in the line edit

    def go_to_ventana(self,ventana):
        window_manager.show_window(ventana)  # Use the window manager to show ventana

    def guardar(self):
        # Save the data from the line edit to the information instance
        self.info.set_data(self.lineEdit2.text())
        print(f"Data saved: {self.info.get_data()}")  # Optional: Print the saved data

def main():
    app = QApplication(sys.argv)
    info = Information()  # Create an instance of the Information class
    global window_manager  # Declare window_manager as global to access it in other classes
    window_manager = WindowManager(info)  # Create an instance of the WindowManager
    window_manager.show_window(Ventana1)  # Show the first window

    try:
        sys.exit(app.exec_())
    except SystemExit:
        print("Exiting")

if __name__ == "__main__":
    main()
