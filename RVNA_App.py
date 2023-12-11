# Importing needed components for application
from PySide6.QtWidgets import QApplication

# sys allows for processing command line arguments
import sys

# imports MainWindow class from separate file
from RVNA_MainWindow import RVNAMainWindow

# subprocess library used to executables
import subprocess

# Opening RVNA.exe external software
RVNA_exe = subprocess.Popen("C:\\VNA\\RVNA\\RVNA.exe")


# Defines python RVNA Application
RVNA_App = QApplication(sys.argv)

# Defines Main Widow Interface for RVNA Application
Main_Window = RVNAMainWindow(RVNA_App)
Main_Window.show()

# Starts event loop - also a blocking function
RVNA_App.exec()

# closes RVNA.exe when python RVNA Application closes
try:
    RVNA_exe.terminate()
except Exception:
    pass
