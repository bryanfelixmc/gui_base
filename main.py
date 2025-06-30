import sys
from PyQt5.uic import loadUi
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QDialog, QApplication

class Information:
    """Class to store and manage information."""
    def __init__(self):
        self.data = ""
    def set_data(self, data):
        self.data = data
    def get_data(self):
        return self.data

class WindowManager:
    """Class to manage transitions between windows."""
    def __init__(self, info):
        self.info = info
        self.current_window = None

    def show_window(self, window_class):
        """Show a new window and hide the current one."""
        if self.current_window is not None:
            self.current_window.hide()  # Hide the current window
        self.current_window = window_class(self.info)  # Create a new window instance
        self.current_window.show()  # Show the new window

class Ventana1(QDialog):
    def __init__(self, info):
        super(Ventana1, self).__init__()
        loadUi("ventana1.ui", self)
        self.info = info  # Store the information instance
        self.next_button.clicked.connect(self.go_to_ventana2)  # Connect the button to the method
    def go_to_ventana2(self):
        window_manager.show_window(Ventana2)  # Use the window manager to show Ventana2

class Ventana2(QDialog):
    def __init__(self, info):
        super(Ventana2, self).__init__()
        loadUi("ventana2.ui", self)
        self.info = info  # Store the information instance
        self.prev_button.clicked.connect(self.ir_ventana1)  # Connect the back button to the method
        self.next_button.clicked.connect(self.ir_ventana3)  # Connect the next button to a new method


    def ir_ventana1(self):
        window_manager.show_window(Ventana1)  # Use the window manager to show Ventana1

    def ir_ventana3(self):
        window_manager.show_window(Ventana3)  # Use the window manager to show Ventana1

class Ventana2_1(QDialog):
    def __init__(self, info):
        super(Ventana2_1, self).__init__()
        loadUi("ventana2.ui", self)
        self.info = info  # Store the information instance
        self.prev_button.clicked.connect(self.ir_ventana1)  # Connect the back button to the method
        self.next_button.clicked.connect(self.ir_ventana3)  # Connect the next button to a new method


    def ir_ventana1(self):
        window_manager.show_window(Ventana1)  # Use the window manager to show Ventana1

    def ir_ventana3(self):
        window_manager.show_window(Ventana3)  # Use the window manager to show Ventana1

class Ventana3(QDialog):
    def __init__(self, info):
        super(Ventana3, self).__init__()
        loadUi("ventana3.ui", self)
        self.info = info  # Store the information instance
        self.prev_button.clicked.connect(self.ir_ventana2)  # Connect the back button to the method
        self.next_button.clicked.connect(self.ir_ventana4)  # Connect the next button to a new method
        self.pushButton3.clicked.connect(self.guardar)  # Connect the save button to the method
        self.lineEdit.setText(self.info.get_data())  # Set the text in the line edit

    def ir_ventana2(self):
        window_manager.show_window(Ventana2)  # Use the window manager to show Ventana1

    def ir_ventana4(self):
        window_manager.show_window(Ventana4)  # Use the window manager to show Ventana1


    def guardar(self):
        # Save the data from the line edit to the information instance
        self.info.set_data(self.lineEdit.text())
        print(f"Data saved: {self.info.get_data()}")  # Optional: Print the saved data

class Ventana4(QDialog):
    def __init__(self, info):
        super(Ventana4, self).__init__()
        loadUi("ventana4.ui", self)
        self.info = info  # Store the information instance
        self.prev_button.clicked.connect(self.ir_ventana3)  # Connect the back button to the method
        self.pushButton3.clicked.connect(self.guardar)  # Connect the save button to the method
        self.lineEdit2.setText(self.info.get_data())  # Set the text in the line edit

    def ir_ventana3(self):
        window_manager.show_window(Ventana3)  # Use the window manager to show Ventana1

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
