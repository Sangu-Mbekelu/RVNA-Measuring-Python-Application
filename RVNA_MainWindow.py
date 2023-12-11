# Imports from python packages

from PySide6.QtWidgets import QMainWindow, QPushButton, QStatusBar, QWidget, QTextEdit, QFrame, QVBoxLayout, QHBoxLayout, QFormLayout, QDialog, QFileDialog, QMessageBox, QLineEdit, QLabel
from PySide6.QtGui import QIcon, QPainter, QFont
from PySide6.QtCore import Signal, QThread, QTimer, QPointF, Qt
from PySide6.QtCharts import QChart, QChartView, QLineSeries, QScatterSeries, QValueAxis
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from datetime import datetime
from numpy import row_stack
from paramiko import SSHClient, AutoAddPolicy
from scp import SCPClient
from os import path, listdir, getcwd, mkdir
import pyvisa
import pandas as pd
import time

# Imports key information from other python file
import User_Pass_Key


class RVNAMainWindow(QMainWindow):

    def __init__(self, app):    # Main Window Constructor
        super().__init__()
        self.app = app
        self.setWindowTitle("RVNA Reading Application")  # Set Window Title
        self.setWindowIcon(QIcon("Resources\\SmithChartIcon.png"))  # Set Window Icon

        # Instance Variables
        self.time_inbetween_measurements = 10  # Default time inbetween measurements
        self.cal_file_directory = "CalFile.cfg"  # Default location of calibration file
        self.local_meas_dir = None
        self.measurement_file_directory = ""
        self.rvna_is_connected = 0
        self.smoothing = 15  # Default measured imaginary impedance smoothing
        self.time_elapsed_min = 0
        self.time_elapsed_max = 30
        self.inflection_frequency_min = 1000
        self.inflection_frequency_max = 1500
        self.frequency_smoothing = 1
        self.init = 1
        self.start_elapsed_time = 0
        # =========================================================================================

        # Log File Path ===========================================================================
        self.log_file_path = str(MeasurementThread.measurements_directory)+"\\0_data_log.txt"
        # =========================================================================================

        # Font used for Graphs ====================================================================
        self.graph_font = QFont()
        self.graph_font.setPointSize(6)
        # =========================================================================================

        # Buttons, Text/Line Edits, and Graphs Layout =============================================
        main_widget = QWidget()  # Define a main widget
        self.setCentralWidget(main_widget)  # The Central Widget includes text editor and start button

        main_frame = QFrame(main_widget)  # Define frame for widgets
        main_frame.setFrameShape(QFrame.Shape.NoFrame)
        # Adding Text Editor as a main way to update users
        #self.main_widget_textedit = QTextEdit(main_frame)
        #self.main_widget_textedit = QTextEdit()
        #self.main_widget_textedit.setReadOnly(1)

        # Adding Push Button to start VNA cal and measurements
        calibrate_measure_button = QPushButton("Calibrate and Start Measurements", main_frame)
        calibrate_measure_button.clicked.connect(self.calibrate_and_start_measurement)

        # Adding Push Button to stop VNA measurements
        stop_button = QPushButton("Stop Measurements", main_frame)
        stop_button.clicked.connect(self.stop_measurement)

        # X-axis used for S11 graph
        self.frequency_axis = QValueAxis()
        self.frequency_axis.setRange(0.85, 4)  # Sets graph from 0.85-4 GHz
        self.frequency_axis.setLabelFormat("%0.2f")
        self.frequency_axis.setLabelsFont(self.graph_font)
        self.frequency_axis.setTickType(QValueAxis.TickType.TicksFixed)
        self.frequency_axis.setTickCount(21)
        self.frequency_axis.setTitleText("Frequency [GHz]")
        # Y-axis used for S11 graph
        self.s11_mag_axis = QValueAxis()
        self.s11_mag_axis.setRange(-40, 0)
        self.s11_mag_axis.setLabelFormat("%0.1f")
        self.s11_mag_axis.setLabelsFont(self.graph_font)
        self.s11_mag_axis.setTickType(QValueAxis.TickType.TicksFixed)
        self.s11_mag_axis.setTickCount(11)
        self.s11_mag_axis.setTitleText("S11 [dB]")

        self.s11_series = QLineSeries()

        # Graph of S11
        self.s11_graph = QChart()
        self.s11_graph.setTitle('Most Recent Antenna Reflection Data')
        self.s11_graph.legend().hide()
        self.s11_graph.addAxis(self.frequency_axis, Qt.AlignmentFlag.AlignBottom)
        self.s11_graph.addAxis(self.s11_mag_axis, Qt.AlignmentFlag.AlignLeft)
        self.s11_graph_view = QChartView(self.s11_graph)
        self.s11_graph_view.setRenderHint(QPainter.RenderHint.Antialiasing)

        # X-axis used for inflection impedance graph
        self.time_elapsed_axis = QValueAxis()
        self.time_elapsed_axis.setRange(self.time_elapsed_min, self.time_elapsed_max)
        self.time_elapsed_axis.setLabelFormat("%0.1f")
        self.time_elapsed_axis.setLabelsFont(self.graph_font)
        self.time_elapsed_axis.setTickType(QValueAxis.TickType.TicksFixed)
        self.time_elapsed_axis.setTickCount(21)
        self.time_elapsed_axis.setTitleText("Time Elapsed [min]")
        # Y-axis used for inflection impedance graph
        self.inflection_frequency_axis = QValueAxis()
        self.inflection_frequency_axis.setRange(self.inflection_frequency_min, self.inflection_frequency_max)
        self.inflection_frequency_axis.setLabelFormat("%0.1f")
        self.inflection_frequency_axis.setLabelsFont(self.graph_font)
        self.inflection_frequency_axis.setTickType(QValueAxis.TickType.TicksFixed)
        self.inflection_frequency_axis.setTickCount(11)
        self.inflection_frequency_axis.setTitleText("Inflection Frequency [MHz]")
        # Y-axis used for inflection impedance graph
        self.s11_min_axis = QValueAxis()
        self.s11_min_axis.setRange(-50, -10)
        self.s11_min_axis.setLabelFormat("%d")
        self.s11_min_axis.setLabelsFont(self.graph_font)
        self.s11_min_axis.setTickType(QValueAxis.TickType.TicksFixed)
        self.s11_min_axis.setTickCount(11)
        self.s11_min_axis.setTitleText("Minimum S11 [dB]")

        self.inflection_frequency_series = QLineSeries()
        self.inflection_frequency_series.setName("Infection Frequency")
        self.s11_min_series = QScatterSeries()
        self.s11_min_series.setName("Minimum S11")
        self.s11_min_series.setMarkerSize(3)
        self.s11_min_series.setBorderColor(Qt.GlobalColor.transparent)

        # Graph of Inflection Impedance
        self.frequency_graph = QChart()
        self.frequency_graph.setTitle('Inflection Frequency Over Time')
        self.frequency_graph.addAxis(self.time_elapsed_axis, Qt.AlignmentFlag.AlignBottom)
        self.frequency_graph.addAxis(self.inflection_frequency_axis, Qt.AlignmentFlag.AlignLeft)
        self.frequency_graph.addAxis(self.s11_min_axis, Qt.AlignmentFlag.AlignRight)
        self.frequency_graph_view = QChartView(self.frequency_graph)
        self.frequency_graph_view.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Line Edits to change graph axis ranges
        self.set_time_elapsed_min = QLineEdit()
        self.set_time_elapsed_min.returnPressed.connect(self.enter_time_elapsed)
        self.set_time_elapsed_max = QLineEdit()
        self.set_time_elapsed_max.returnPressed.connect(self.enter_time_elapsed)
        self.set_inflection_frequency_min = QLineEdit()
        self.set_inflection_frequency_min.returnPressed.connect(self.enter_inflection_frequency)
        self.set_inflection_frequency_max = QLineEdit()
        self.set_inflection_frequency_max.returnPressed.connect(self.enter_inflection_frequency)

        # Line Edit for smoothing
        self.smoothing_label = QLabel()
        self.smoothing_label.setFixedWidth(60)
        self.smoothing_label.setText("Smoothing: ")
        self.set_smoothing = QLineEdit()
        self.set_smoothing.setFixedWidth(50)
        self.set_smoothing.returnPressed.connect(self.enter_smoothing)

        # Two Form Layouts, combined with a Horizontal Layout for the graph changing Line Edits
        time_elapsed_changes_layout = QFormLayout()
        time_elapsed_changes_layout.addRow("Time Elapsed (min): ", self.set_time_elapsed_min)
        time_elapsed_changes_layout.addRow("Time Elapsed (max): ", self.set_time_elapsed_max)
        inflection_impedance_changes_layout = QFormLayout()
        inflection_impedance_changes_layout.addRow("Inflection Frequency (min): ", self.set_inflection_frequency_min)
        inflection_impedance_changes_layout.addRow("Inflection Frequency (max): ", self.set_inflection_frequency_max)
        graphing_changes_layout = QHBoxLayout()
        graphing_changes_layout.addLayout(time_elapsed_changes_layout)
        graphing_changes_layout.addLayout(inflection_impedance_changes_layout)

        # Vertical Layout for Two Graphs
        graph_layout = QVBoxLayout()
        graph_layout.addWidget(self.frequency_graph_view)
        graph_layout.addWidget(self.s11_graph_view)

        # Horizontal Layout for buttons
        button_layout = QHBoxLayout()
        button_layout.addWidget(calibrate_measure_button)
        button_layout.addWidget(stop_button)
        button_layout.addWidget(self.smoothing_label)
        button_layout.addWidget(self.set_smoothing)

        # Vertical Layout to display Central Widgets
        central_layout = QVBoxLayout(main_frame)
        #central_layout.addWidget(self.main_widget_textedit)
        central_layout.addLayout(graphing_changes_layout)
        central_layout.addLayout(button_layout)
        #central_layout.setStretchFactor(main_frame, 1)
        main_frame.setLayout(central_layout)  # Sets layout in main_frame

        # To ensure central widget fits frame
        central_widget_layout = QVBoxLayout(main_widget)
        central_widget_layout.addWidget(main_frame)
        central_widget_layout.addLayout(graph_layout)
        central_widget_layout.setStretchFactor(graph_layout, 1)
        main_widget.setLayout(central_widget_layout)  # Sets the frame in the main widget
        # =========================================================================================

        # Timer Initialization for Measurement Thread =============================================
        # Initializing the Thread used to Start the Measurements
        self.measurement = MeasurementThread(self.smoothing)
        self.measurement.measurement_update.connect(self.measurement_update_event)
        self.measurement.measurements_filedirectory.connect(self.get_measurement_file)
        self.measurement.finished.connect(self.graphing)  # graph function runs everytime

        # Timer Initialized to Start Measurement Thread
        self.measurement_timer = QTimer()
        self.measurement_timer.timeout.connect(self.run_measurement)
        self.measurement_timer.setInterval(int(self.time_inbetween_measurements) * 1000)
        # =========================================================================================

        # Initializing Timer and File Transfer Thread =============================================
        self.data_log_transfer = ServerTransferThread("data_log")
        self.s_parameter_transfer = ServerTransferThread("s_parameters")

        self.data_log_transfer_timer = QTimer()
        self.data_log_transfer_timer.timeout.connect(self.data_file_transfer)
        self.data_log_transfer_timer.setInterval(3000)  # Every 3 seconds, the program will attempt to connect to the server, transfer files, then close the connection
        self.data_log_transfer_timer.start()

        self.s_parameters_transfer_timer = QTimer()
        self.s_parameters_transfer_timer.timeout.connect(self.s_parameter_file_transfer)
        self.s_parameters_transfer_timer.setInterval(60000)  # Every 60 seconds, the program will attempt to connect to the server, transfer files, then close the connection
        self.s_parameters_transfer_timer.start()
        # =========================================================================================

        # Menubar =================================================================================
        self.menu_bar = self.menuBar()
        # Settings Menu (Used to change default settings)
        settings_menu = self.menu_bar.addMenu("Settings")
        # Time Change Action allows user to change the time period inbetween VNA measurements
        time_change_action = settings_menu.addAction("Time Inbetween Measurements")
        self.time_change_window = TimeChangeWidget()
        self.time_change_window.submit_time.connect(self.time_change)  # Connecting TimeChangeWidget signal to the Main Window slot
        time_change_action.triggered.connect(self.time_change_window.show)
        # Calibration State File Location Action allows user to change the directory that the Cal State File is located in
        cal_state_file_location_action = settings_menu.addAction("Calibration State Location")
        cal_state_file_location_action.triggered.connect(self.cal_state_location)
        # Change Smoothing Window Action allows user to change inflection impedance / frequency smoothing window during measurement
        change_smoothing_action = settings_menu.addAction("Smoothing Window")
        self.change_smoothing_window = SmoothingChangeWidget()
        self.change_smoothing_window.impedance_smoothing.connect(self.smoothing_change)
        change_smoothing_action.triggered.connect(self.change_smoothing_window.show)
        # Help Menu (Used to help users)
        help_menu = self.menu_bar.addMenu("Help")
        pdf_help_action = help_menu.addAction("Help Document")
        self.pdf_view_window = HelpWidget()
        pdf_help_action.triggered.connect(self.pdf_view_window.show)
        # =========================================================================================

        # Status Bar ==============================================================================
        self.setStatusBar(QStatusBar(self))

        # Maximize Window
        self.showMaximized()  # Setting Fullscreen

    def calibrate_and_start_measurement(self):
        global CMT
        # RVNA Software Connection =====================================
        if self.rvna_is_connected == 0:
            rm = pyvisa.ResourceManager('@py')  # use pyvisa-py as backend

            try:
                CMT  # Used to check if there has been an object created to connect to the RVNA
            except NameError:
                try:
                    CMT = rm.open_resource('TCPIPO::127.0.0.1::5025::SOCKET')  # Connects to RVNA application using SCPI

                    connection_message = "Connected to VNA\n"
                    #self.main_widget_textedit.append(connection_message)  # Updates Text Editor
                    self.rvna_is_connected = 1
                except:
                    error_message = "Failed to Connect to VNA\nCheck RVNA Connection to Laptop\n"
                    #self.main_widget_textedit.append(error_message)  # Updates Text Editor
                    return

        CMT.read_termination = '\n'  # The VNA ends each line with this. Reads will time out without this
        CMT.timeout = 10000  # Set longer timeout period for slower sweeps

        # RVNA Calibration Process ======================================
        CMT.write(f"MMEM:LOAD:STAT {self.cal_file_directory}")  # Recalls calibration state with specified file
        CMT.write("DISP:WIND:SPL 2")  # Allocate 2 trace windows
        CMT.write("CALC1:PAR:COUN 3")  # 3 Traces
        CMT.write("CALC1:PAR1:DEF S11")  # Choose S11 for trace 1
        CMT.write("CALC1:PAR2:DEF S11")  # Choose S11 for trace 2
        CMT.write("CALC1:PAR3:DEF S11")  # Choose S11 for trace 3

        CMT.write("CALC1:PAR1:SEL")  # Selects Trace 1 and Phase Format
        CMT.write("CALC1:FORM PHAS")

        CMT.write("CALC1:PAR2:SEL")  # Selects Trace 2 and Smith Chart Format
        CMT.write("CALC1:FORM SMIT")

        CMT.write("CALC1:PAR3:SEL")  # Selects Trace 3 and Log Mag Format
        CMT.write("CALC1:FORM MLOG")

        CMT.query("*OPC?")  # Wait for measurement to complete

        # open calibration window
        cal_window = CalibrationDialog()
        cal_accepted = cal_window.exec()

        if cal_accepted == 0:
            user_alert = QMessageBox()
            user_alert.setWindowTitle("Start Measurement Failed")
            user_alert.setText("Calibration Not Accepted")
            user_alert.setInformativeText("Restart")
            user_alert.setIcon(QMessageBox.Icon.Critical)
            user_alert.exec()
            return
        # ==========================================================================================================

        # Get Measurements Folder Name ======================================
        self.receive_directory()

        if self.local_meas_dir is None:
            user_alert = QMessageBox()
            user_alert.setWindowTitle("Start Measurement Failed")
            user_alert.setText("Did Not Enter a Valid Folder Name")
            user_alert.setInformativeText("Restart")
            user_alert.setIcon(QMessageBox.Icon.Critical)
            user_alert.exec()
            return

        MeasurementThread.measurements_directory = path.normpath(self.local_meas_dir)  # Changes MeasurementThread class variable
        ServerTransferThread.measurements_directory = path.normpath(self.local_meas_dir)  # Changes ServerTransferThread class variable

        #self.main_widget_textedit.append("RVNA is calibrated\n")  # Updates Text Editor

        self.measurement_timer.start()

        self.menu_bar.hide()

        self.statusBar().showMessage("RVNA calibrated", 10000)  # Updates Status Bar

    def run_measurement(self):
        self.measurement.start()

    def data_file_transfer(self):
        self.data_log_transfer.start()

    def s_parameter_file_transfer(self):
        self.s_parameter_transfer.start()

    def measurement_update_event(self, meas_update):  # Takes signal from measurement Thread
        #self.main_widget_textedit.append(meas_update)  # Updates Text Editor
        pass

    def get_measurement_file(self, file):
        self.measurement_file_directory = MeasurementThread.measurements_directory+"\\"+file[0]
        self.log_file_path = MeasurementThread.measurements_directory+file[1]

    def graphing(self):  # Is called after measurement thread finishes
        self.s11_series.clear()  # Clears data from series
        self.s11_graph.removeSeries(self.s11_series)  # Removes series from graph
        current_file_contents = pd.read_csv(self.measurement_file_directory)  # Reads measurement file as dataframe
        frequency = current_file_contents['Frequency [Hz]'].tolist()  # Creates frequency list
        s11_mag = current_file_contents['S11 [dB]'].tolist()  # Creates S11 magnitude list

        for i in range(len(frequency)):
            self.s11_series.append(QPointF(frequency[i] / 1e9, s11_mag[i]))  # Appends all points to series

        self.s11_graph.addSeries(self.s11_series)  # Adds series to graph
        self.s11_series.attachAxis(self.frequency_axis)  # Attaches both axis to the series
        self.s11_series.attachAxis(self.s11_mag_axis)
        self.s11_graph.setTitle('Most Recent Antenna Reflection Data: Resonating at %0.2f MHz' % (float(current_file_contents['Inflection Frequency [Hz]'][0]) / 1e6))  # Changes title based on recent inflection impedance value

        self.inflection_frequency_series.clear()  # Clears data from series
        self.s11_min_series.clear()  # Clears data from series
        self.frequency_graph.removeSeries(self.inflection_frequency_series)  # Removes series from graph
        self.frequency_graph.removeSeries(self.s11_min_series)
        log_file_contents = pd.read_csv(self.log_file_path)  # Reads data log file as dataframe
        inflection_frequency = log_file_contents['Inflection Frequency [Hz]'].rolling(self.frequency_smoothing).mean().tolist()  # Creates inflection frequency list
        inflection_frequency = inflection_frequency[(self.frequency_smoothing - 1):]
        elapsed_time_seconds = log_file_contents['Elapsed Times [s]'].tolist()  # Creates elapsed time list
        min_s11 = log_file_contents['S11 at Inflection Frequency [dB]'].tolist()  # Creates minimum S11 list

        for i in range(len(inflection_frequency)):
            self.inflection_frequency_series.append(QPointF((elapsed_time_seconds[i + (self.frequency_smoothing - 1)] / 60), (inflection_frequency[i]) / 1e6))  # Appends all points to series

        for i in range(len(elapsed_time_seconds)):
            self.s11_min_series.append(QPointF((elapsed_time_seconds[i] / 60), min_s11[i]))

        self.frequency_graph.addSeries(self.inflection_frequency_series)  # Adds series to graph
        self.frequency_graph.addSeries(self.s11_min_series)
        self.inflection_frequency_series.attachAxis(self.inflection_frequency_axis)  # Attaches both axis to the inflection frequency series
        self.inflection_frequency_series.attachAxis(self.time_elapsed_axis)
        self.s11_min_series.attachAxis(self.s11_min_axis)  # Attaches both axis to the minimum S11 series
        self.s11_min_series.attachAxis(self.time_elapsed_axis)

    def stop_measurement(self):
        self.menu_bar.show()
        self.measurement_timer.stop()
        #self.main_widget_textedit.append("Measurements Stopped")  # Updates Text Editor

    def enter_time_elapsed(self):
        min_time = self.set_time_elapsed_min.text()
        max_time = self.set_time_elapsed_max.text()
        if min_time != "" and max_time == "":
            try:
                if float(min_time) < float(self.time_elapsed_max):
                    self.time_elapsed_min = float(min_time)
                    self.time_elapsed_axis.setRange(self.time_elapsed_min, self.time_elapsed_max)
                else:
                    pass
            except ValueError:
                pass
        elif min_time == "" and max_time != "":
            try:
                if float(max_time) > float(self.time_elapsed_min):
                    self.time_elapsed_max = float(max_time)
                    self.time_elapsed_axis.setRange(self.time_elapsed_min, self.time_elapsed_max)
                else:
                    pass
            except ValueError:
                pass
        else:
            try:
                if float(max_time) > float(min_time):
                    self.time_elapsed_max = float(max_time)
                    self.time_elapsed_min = float(min_time)
                    self.time_elapsed_axis.setRange(self.time_elapsed_min, self.time_elapsed_max)
                else:
                    pass
            except ValueError:
                pass

    def enter_inflection_frequency(self):
        min_imp = self.set_inflection_frequency_min.text()
        max_imp = self.set_inflection_frequency_max.text()
        if min_imp != "" and max_imp == "":
            try:
                if float(min_imp) < float(self.inflection_frequency_max):
                    self.inflection_frequency_min = float(min_imp)
                    self.inflection_frequency_axis.setRange(self.inflection_frequency_min, self.inflection_frequency_max)
                else:
                    pass
            except ValueError:
                pass
        elif min_imp == "" and max_imp != "":
            try:
                if float(max_imp) > float(self.inflection_frequency_min):
                    self.inflection_frequency_max = float(max_imp)
                    self.inflection_frequency_axis.setRange(self.inflection_frequency_min, self.inflection_frequency_max)
                else:
                    pass
            except ValueError:
                pass
        else:
            try:
                if float(max_imp) > float(min_imp):
                    self.inflection_frequency_max = float(max_imp)
                    self.inflection_frequency_min = float(min_imp)
                    self.inflection_frequency_axis.setRange(self.inflection_frequency_min, self.inflection_frequency_max)
                else:
                    pass
            except ValueError:
                pass

    def enter_smoothing(self):
        smoothing = self.set_smoothing.text()
        if smoothing != "":
            try:
                int(smoothing)
                self.frequency_smoothing = int(smoothing)
            except ValueError:
                pass

    def time_change(self, time_inbetween):
        self.measurement_timer.setInterval(int(time_inbetween) * 1000)
        #self.main_widget_textedit.append(f"Time inbetween Measurements Changed to {time_inbetween} Seconds")
        self.statusBar().showMessage(f"Time inbetween Measurements Changed to {time_inbetween} Seconds", 10000)

    def smoothing_change(self, smoothing):
        MeasurementThread.input_imaginary_impedance_smoothing_window = smoothing
        #self.main_widget_textedit.append(f"Imaginary Impedance Smoothing Changed to {smoothing}")
        self.statusBar().showMessage(f"Imaginary Impedance Smoothing Changed to {smoothing}", 10000)

    def cal_state_location(self):
        self.cal_file_directory = QFileDialog().getOpenFileName(parent=self)[0]
        if self.cal_file_directory == "":
            return
        #self.main_widget_textedit.append(f"Cal State File Location Changed to {self.cal_file_directory}")
        self.statusBar().showMessage(f"Cal State File Location Changed to {self.cal_file_directory}", 10000)

    def receive_directory(self):
        def get_dir(folder_name):
            self.local_meas_dir = getcwd() + "\\Measurement_Data\\" + folder_name

        directory = FolderNameDialog()
        directory.folder_name.connect(get_dir)
        directory.exec()


class MeasurementThread(QThread):
    global CMT
    # Signal emitted in run function
    measurement_update = Signal(str)
    measurements_filedirectory = Signal(list)

    # Initialized class variables
    input_imaginary_impedance_smoothing_window = None
    measurements_directory = None

    def __init__(self, smoothing_variable):
        MeasurementThread.input_imaginary_impedance_smoothing_window = smoothing_variable
        self.log_list = []
        self.init = 1
        self.start_elapsed_time = 0.0
        self.numb_file = 1
        super().__init__()

    def run(self):
        CMT.write("TRIG:SOUR BUS")  # Set sweep source to BUS for automated measurement
        CMT.query("*OPC?")  # Wait for measurement to complete

        CMT.write("TRIG:SING")  # Trigger a single sweep
        CMT.query("*OPC?")  # Wait for measurement to complete

        current_datetime = datetime.now()

        if self.init == 1:
            self.start_elapsed_time = time.time()
            end_elapsed_time = self.start_elapsed_time
        else:
            end_elapsed_time = time.time()

        current_time_hour = current_datetime.strftime("%H")  # Logging the current hour
        current_time_minute = current_datetime.strftime("%M")  # Logging the current minute
        current_time_second = current_datetime.strftime("%S")  # Logging the current second

        # Read frequency data
        freq = CMT.query_ascii_values("SENS1:FREQ:DATA?")

        # Read smith chart impedance data
        CMT.write("CALC1:PAR2:SEL")
        imp = CMT.query_ascii_values("CALC1:DATA:FDAT?")  # Get data as string
        real_imp = imp[::2]
        imag_imp = imp[1::2]

        # Read log mag data
        CMT.write("CALC1:PAR3:SEL")
        log_mag = CMT.query_ascii_values("CALC1:DATA:FDAT?")  # Get data as string
        log_mag = log_mag[::2]

        # Read phase data
        CMT.write("CALC1:PAR1:SEL")
        phase = CMT.query_ascii_values("CALC1:DATA:FDAT?")  # Get data as string
        phase = phase[::2]

        vna_temp = CMT.write("SYST:TEMP:SENS<1>?")

        CMT.write("TRIG:SOUR INT")  # Set sweep source to INT after measurements are done

        CMT.query("*OPC?")  # Wait for measurement to complete

        current_time_hour_list = [int(current_time_hour)] * len(freq)  # Creates lists from single values of hour, minute, and second

        current_time_minute_list = [int(current_time_minute)] * len(freq)

        current_time_second_list = [int(current_time_second)] * len(freq)

        vna_temp = [(vna_temp * (9 / 5)) + 32] * len(freq)  # Creates list from VNA temp value

        self.measurement_update.emit(f"Measurement Taken at {current_datetime.strftime('%m-%d-%Y_%H-%M-%S')}\n")  # Emits signal of the time a measurement was taken to the TextEdit

        # Creates dictionary that will be in dataframe object
        data_dictionary = {'Current Hour': current_time_hour_list, 'Current Minute': current_time_minute_list,
                           'Current Second': current_time_second_list, 'Frequency [Hz]': freq, 'S11 [dB]': log_mag,
                           'S11 Phase [DEG]': phase, 'Zin [RE ohm]': real_imp, 'Zin [IM ohm]': imag_imp,
                           'VNA Temp [F]': vna_temp}

        data_frame = pd.DataFrame(data_dictionary)  # Creating dataframe

        # Initializations for minimum S11 and inflection impedance
        returnloss_mag_min = 0.0
        inflection_impedance = 0.0
        inflection_frequency = 0.0

        imag_imp_df = pd.DataFrame({'Impedance': imag_imp}) # Imaginary impedance list

        corresponding_returnloss_mag = log_mag
        corresponding_impedance_real = real_imp
        # Creates a rolling average of imaginary impedance values
        smoothed_input_impedance_imag = imag_imp_df['Impedance'].rolling(MeasurementThread.input_imaginary_impedance_smoothing_window).mean()
        # Finds when the imagnary impedance crosses the zero to get closes to a purely real impedance
        for i in range(1, len(smoothed_input_impedance_imag)-1):  # Calculates the inflection frequency, real inflection impedance, and minimum S11. The inflection impedance is found by searching the measured close-to-purely real impedances and defining the one with the lowest magnitude S11 as the inflection impedance
            if abs(smoothed_input_impedance_imag[i - 1]) > abs(smoothed_input_impedance_imag[i]) and abs(smoothed_input_impedance_imag[i + 1]) > abs(smoothed_input_impedance_imag[i]):
                if corresponding_returnloss_mag[i] < returnloss_mag_min:    # Inflection impedance will be the purely real impedance with the lowest magnitude S11
                    returnloss_mag_min = corresponding_returnloss_mag[i]
                    inflection_frequency = freq[i]
                    inflection_impedance = corresponding_impedance_real[i]

        # Creating lists from single values
        real_inflection_impedance = [inflection_impedance] * len(freq)
        min_returnloss = [returnloss_mag_min] * len(freq)
        min_returnloss_freq = [inflection_frequency] * len(freq)

        # Inserts extra columns of data
        data_frame.insert(3, 'Inflection Frequency [Hz]', min_returnloss_freq)
        data_frame.insert(9, 'S11 at Inflection Frequency [dB]', min_returnloss)
        data_frame.insert(10, 'Inflection Impedance [RE ohm]', real_inflection_impedance)

        file_name = f'{self.numb_file}_' + 'S_parameters_' + str(current_datetime.strftime('%m-%d-%Y_%H-%M-%S')) + '.txt'  # Creates file name based on time measurement was taken

        self.numb_file += 1  # Increment by 1, makes listing s-parameter files by name while maintaining proper order easier

        data_frame.to_csv(MeasurementThread.measurements_directory+"\\"+file_name, index=False, sep=',', header=True)  # Saves dataframe as csv

        elapsed_time_seconds = round(abs(end_elapsed_time - self.start_elapsed_time))  # Calculating elapsed time from a start and end time

        # Creates data for data log file
        log_new_row = [int(current_time_hour), int(current_time_minute), int(current_time_second), elapsed_time_seconds, inflection_frequency, inflection_impedance, returnloss_mag_min]

        if self.init == 1:
            self.log_list = log_new_row  # If it is the first measurement, equate list as the first row
            self.init = 0
            log_frame_dict = {'Current Hour': [self.log_list[0]], 'Current Minute': [self.log_list[1]], 'Current Second': [self.log_list[2]], 'Elapsed Times [s]': [self.log_list[3]], 'Inflection Frequency [Hz]': [self.log_list[4]], 'Inflection Impedance [RE ohm]': [self.log_list[5]], 'S11 at Inflection Frequency [dB]': [self.log_list[6]]}  # Initializes dictionary for first version of dataframe
            log_frame = pd.DataFrame(log_frame_dict)
        elif self.numb_file == 3:  # flips back initial vertical array for dataframe creation
            self.log_list = row_stack((self.log_list, log_new_row))  # Appends row to existing list from the Main Window
            self.log_list.tolist()
            log_frame = pd.DataFrame(self.log_list, columns=['Current Hour', 'Current Minute', 'Current Second', 'Elapsed Times [s]', 'Inflection Frequency [Hz]', 'Inflection Impedance [RE ohm]', 'S11 at Inflection Frequency [dB]'])  # Initializes dataframe
        else:
            self.log_list = row_stack((self.log_list, log_new_row))  # Appends row to existing list from the Main Window
            self.log_list.tolist()
            log_frame = pd.DataFrame(self.log_list, columns=['Current Hour', 'Current Minute', 'Current Second', 'Elapsed Times [s]', 'Inflection Frequency [Hz]', 'Inflection Impedance [RE ohm]', 'S11 at Inflection Frequency [dB]'])  # Initializes dataframe

        log_frame.to_csv(MeasurementThread.measurements_directory+"\\0_data_log.txt", index=False, sep=',', header=True)  # Saves dataframe as csv

        self.measurements_filedirectory.emit([file_name, "\\0_data_log.txt"])


class ServerTransferThread(QThread):
    measurements_directory = None

    def __init__(self, data_type):
        super().__init__()
        # Determines wither the thread object with transmit the data_log file or the s-parameters
        self.type_of_data_transfer = data_type
        # Setting Constant Variables for SSH
        self.ssh = SSHClient()  # Defines SSH client
        self.ssh.set_missing_host_key_policy(AutoAddPolicy())  # Adds host key if missing
        self.sftp_session = None
        self.scp = None

        # Server Access Information
        self.server_host = User_Pass_Key.hostname
        self.server_user = User_Pass_Key.user
        self.server_password = User_Pass_Key.password
        self.server_root_directory = User_Pass_Key.remote_path

        # Used to increment through list
        self.numb_file = 1

        self.connection_var = 0

    def run(self):
        if ServerTransferThread.measurements_directory is not None:
            if self.connection_var == 0:
                try:
                    self.ssh.connect(self.server_host, username=self.server_user, password=self.server_password)  # Establishes SSH connection
                    self.sftp_session = self.ssh.open_sftp()  # Opens SFTP session
                    self.scp = SCPClient(self.ssh.get_transport())
                    self.connection_var = 1
                except:
                    return
            else:
                pass

            user_named_folder = path.basename(path.normpath(ServerTransferThread.measurements_directory))  # Gets name of user specified directory

            try:
                self.sftp_session.chdir(User_Pass_Key.remote_path + user_named_folder)  # Changes directory to specified file on the server
            except:
                try:
                    self.sftp_session.mkdir(User_Pass_Key.remote_path + user_named_folder)  # Creates directory of specified file on server
                    self.sftp_session.chdir(User_Pass_Key.remote_path + user_named_folder)  # Changes directory to specified file on the server
                except:
                    self.connection_var = 0
                    return

            s_parameter_list = listdir(ServerTransferThread.measurements_directory)  # get all files in a directory
            s_parameter_list_full_path = [path.join(ServerTransferThread.measurements_directory, f) for f in s_parameter_list]  # add absolute paths to each file
            s_parameter_list_full_path.sort(key=lambda x: path.getmtime(x))  # sorts paths of files by date created
            s_parameter_list = [path.basename(g) for g in s_parameter_list_full_path]  # Gets list of files with just the name

            if self.type_of_data_transfer == "data_log":
                try:
                    if len(s_parameter_list) > self.numb_file:
                        try:
                            self.sftp_session.stat(self.sftp_session.getcwd() + '/' + "0_data_log.txt")
                            self.sftp_session.chmod(self.sftp_session.getcwd() + '/' + "0_data_log.txt", 0o666)
                            self.sftp_session.chmod(self.sftp_session.getcwd() + '/' + "Latest_Sparams.txt", 0o666)
                        except:
                            pass
                        self.scp.put(ServerTransferThread.measurements_directory + '\\' + "0_data_log.txt", self.sftp_session.getcwd() + '/' + "0_data_log.txt")  # Copies new data log from local to remote server
                        self.scp.put(ServerTransferThread.measurements_directory + '\\' + s_parameter_list[-2], self.sftp_session.getcwd() + '/' + "Latest_Sparams.txt")  # Copies latest s-parameter file to
                    else:
                        pass
                except:
                    self.connection_var = 0
                    pass

            elif self.type_of_data_transfer == "s_parameters":

                files_in_server = self.sftp_session.listdir()
                files_in_server_full_paths = [self.sftp_session.getcwd() + '/' + x for x in files_in_server]

                [self.sftp_session.chmod(x, 0o444) for x in files_in_server_full_paths]
                try:
                    [self.s_parameter_file_put(x, y) for x, y in zip(s_parameter_list_full_path, s_parameter_list)]  # Copies new data log from local to remote server
                    [self.sftp_session.chmod(x, 0o666) for x in files_in_server_full_paths]

                except:
                    self.connection_var = 0
        else:
            pass

    def s_parameter_file_put(self, x, y):
        try:
            self.scp.put(x, self.sftp_session.getcwd() + '/' + y)
        except:
            pass


class CalibrationDialog(QDialog):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Calibration Check")  # Set Window Title
        self.setWindowIcon(QIcon("Resources\\SmithChartIcon.png"))  # Set Window Icon
        self.resize(1250, 600)  # Setting Window Size

        # setting default cal state. 1 = cal free space, 2 = cal on body
        self.cal_state = 1

        # Font ================================================================
        self.text_font = QFont()
        self.text_font.setPointSize(15)
        self.button_font = QFont()
        self.button_font.setPointSize(10)
        self.graph_font = QFont()
        self.graph_font.setPointSize(10)
        # S11 Graph ===========================================================
        # X-axis used for S11 graph
        self.frequency_axis = QValueAxis()
        self.frequency_axis.setRange(0.85, 4)  # Sets graph from 0.85-4 GHz
        self.frequency_axis.setLabelFormat("%0.2f")
        self.frequency_axis.setLabelsFont(self.graph_font)
        self.frequency_axis.setTickType(QValueAxis.TickType.TicksFixed)
        self.frequency_axis.setTickCount(21)
        self.frequency_axis.setTitleText("Frequency [GHz]")
        # Y-axis used for S11 graph
        self.s11_mag_axis = QValueAxis()
        self.s11_mag_axis.setRange(-40, 0)
        self.s11_mag_axis.setLabelFormat("%0.1f")
        self.s11_mag_axis.setLabelsFont(self.graph_font)
        self.s11_mag_axis.setTickType(QValueAxis.TickType.TicksFixed)
        self.s11_mag_axis.setTickCount(11)
        self.s11_mag_axis.setTitleText("S11 [dB]")

        self.s11_series = QLineSeries()

        # Graph of S11
        self.s11_graph = QChart()
        self.s11_graph.setTitle('Antenna Return Loss')
        self.s11_graph.setTitleFont(self.text_font)
        self.s11_graph.legend().hide()
        self.s11_graph.addAxis(self.frequency_axis, Qt.AlignmentFlag.AlignBottom)
        self.s11_graph.addAxis(self.s11_mag_axis, Qt.AlignmentFlag.AlignLeft)
        self.s11_graph_view = QChartView(self.s11_graph)
        self.s11_graph_view.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Text Prompts and Buttons ============================================
        # Adding Text Prompt for User
        self.cal_prompt = QLabel()
        self.cal_prompt.setText("Does the Antenna Resonate in Free Space?")
        self.cal_prompt.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cal_prompt.setFont(self.text_font)

        # Adding Push Button to continue cal
        accept_cal_button = QPushButton("Accept Calibration")
        accept_cal_button.clicked.connect(self.continue_cal)
        accept_cal_button.setFont(self.button_font)

        # Adding Push Button to exit cal
        exit_cal_button = QPushButton("Exit Calibration")
        exit_cal_button.clicked.connect(self.exit_cal)
        exit_cal_button.setFont(self.button_font)

        # Layout ==============================================================
        button_layout = QHBoxLayout()
        button_layout.addWidget(accept_cal_button)
        button_layout.addWidget(exit_cal_button)

        cal_layout = QVBoxLayout()
        cal_layout.addWidget(self.s11_graph_view)
        cal_layout.addWidget(self.cal_prompt)
        cal_layout.addLayout(button_layout)
        self.setLayout(cal_layout)

        # S11 Graph Update Timer ==============================================
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.graphing)
        self.update_timer.setInterval(200)  # Every 1/4 second, the program will cal graph
        self.update_timer.start()

    def graphing(self):
        global CMT
        CMT.write("TRIG:SOUR BUS")  # Set sweep source to BUS for automated measurement
        CMT.query("*OPC?")  # Wait for measurement to complete

        CMT.write("TRIG:SING")  # Trigger a single sweep
        CMT.query("*OPC?")  # Wait for measurement to complete

        # Read frequency data
        frequency = CMT.query_ascii_values("SENS1:FREQ:DATA?")

        # Read log mag data
        CMT.write("CALC1:PAR3:SEL")
        log_mag = CMT.query_ascii_values("CALC1:DATA:FDAT?")  # Get data as string
        s11_mag = log_mag[::2]

        self.s11_series.clear()  # Clears data from series
        self.s11_graph.removeSeries(self.s11_series)  # Removes series from graph

        for i in range(len(frequency)):
            self.s11_series.append(QPointF(frequency[i]/1e9, s11_mag[i]))  # Appends all points to series

        self.s11_graph.addSeries(self.s11_series)  # Adds series to graph
        self.s11_series.attachAxis(self.frequency_axis)  # Attaches both axis to the series
        self.s11_series.attachAxis(self.s11_mag_axis)

    def continue_cal(self):
        if self.cal_state == 1:
            user_alert = QMessageBox()
            user_alert.setWindowTitle("Attach Antenna")
            user_alert.setText("Attach the Antenna onto the Body")
            user_alert.setInformativeText("Close this window and use the graph to ensure the antenna resonates while attached to the body")
            user_alert.setIcon(QMessageBox.Icon.Information)
            user_alert.exec()
            self.cal_prompt.setText("Does the Antenna Resonate on the Body?")
            self.cal_state += 1  # advance cal state
        elif self.cal_state == 2:
            self.update_timer.stop()
            self.accept()  # user has determined calibration is good

    def exit_cal(self):
        self.update_timer.stop()
        self.reject()


class FolderNameDialog(QDialog):    # Window used so user can input folder name
    folder_name = Signal(str)  # Signal that will be emitted to Main Window Object

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Enter Folder Name")  # Set Window Title
        self.setWindowIcon(QIcon("Resources\\FileExplorerIcon.png"))
        self.resize(400, 100)  # Set Window Size

        # The two labels and text editor used to convey the information user must input
        text_editor_label = QLabel("Folder Name:")
        self.line_edit = QLineEdit()

        # Initial Horizontal layout used to order the labels and editor
        text_edit_layout = QHBoxLayout()
        text_edit_layout.addWidget(text_editor_label)
        text_edit_layout.addWidget(self.line_edit)

        # Adding the button to receive the data input by user
        set_name_button = QPushButton("Ok")
        set_name_button.clicked.connect(self.set_folder_name)

        # Vertical layout used to place button below text editor
        full_layout = QVBoxLayout()
        full_layout.addLayout(text_edit_layout)
        full_layout.addWidget(set_name_button)

        # Sets window layout
        self.setLayout(full_layout)

    def set_folder_name(self):
        if path.exists("Measurement_Data\\" + self.line_edit.text()):
            string_error = QMessageBox()
            string_error.setWindowTitle("Error")
            string_error.setText("Folder Already Exists")
            string_error.setIcon(QMessageBox.Icon.Critical)
            string_error.setDefaultButton(QMessageBox.StandardButton.Ok)
            string_error.exec()
        elif " " in self.line_edit.text():
            string_error = QMessageBox()
            string_error.setWindowTitle("Error")
            string_error.setText("Please Avoid Spaces")
            string_error.setIcon(QMessageBox.Icon.Critical)
            string_error.setDefaultButton(QMessageBox.StandardButton.Ok)
            string_error.exec()
        else:
            try:
                mkdir("Measurement_Data\\" + self.line_edit.text())
                self.folder_name.emit(self.line_edit.text())  # Signal is emitted to Main Window
                self.close()
            except Exception:
                string_error = QMessageBox()
                string_error.setWindowTitle("Error")
                string_error.setText("Invalid Folder Character Name")
                string_error.setInformativeText("{/, \\, <, >, :, \", |, ?, *}")
                string_error.setIcon(QMessageBox.Icon.Critical)
                string_error.setDefaultButton(QMessageBox.StandardButton.Ok)
                string_error.exec()


class TimeChangeWidget(QWidget):    # Window used to change time inbetween measurements
    submit_time = Signal(str)  # Signal that will be emitted to Main Window Object

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Change Measurement Time")  # Set Window Title
        self.setWindowIcon(QIcon("Resources\\ClockIcon.png"))
        self.resize(400, 100)  # Set Window Size

        # The two labels and text editor used to convey the information user must input
        text_editor_label = QLabel("Time Inbetween Measurements:")
        units_label = QLabel("(sec)")
        self.line_edit = QLineEdit()

        # Initial Horizontal layout used to order the labels and editor
        text_edit_layout = QHBoxLayout()
        text_edit_layout.addWidget(text_editor_label)
        text_edit_layout.addWidget(self.line_edit)
        text_edit_layout.addWidget(units_label)

        # Adding the button to receive the data input by user
        set_time_button = QPushButton("Set Time")
        set_time_button.clicked.connect(self.set_time)

        # Vertical layout used to place button below text editor
        full_layout = QVBoxLayout()
        full_layout.addLayout(text_edit_layout)
        full_layout.addWidget(set_time_button)

        # Sets window layout
        self.setLayout(full_layout)

    def set_time(self):
        try:
            int(self.line_edit.text())
            self.submit_time.emit(self.line_edit.text())  # Signal is emitted to Main Window
            self.close()
        except ValueError:
            string_error = QMessageBox()
            string_error.setWindowTitle("Error")
            string_error.setText("Only Use Digits")
            string_error.setIcon(QMessageBox.Icon.Critical)
            string_error.setDefaultButton(QMessageBox.StandardButton.Ok)
            string_error.exec()


class SmoothingChangeWidget(QWidget):    # Window used to change calibration file location
    impedance_smoothing = Signal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Change Impedance Smoothing")  # Set Window Title
        self.setWindowIcon(QIcon("Resources\\PlotIcon.png"))
        self.resize(400, 100)  # Set Window Size

        # The label and text editor used to convey the information user must input
        text_editor_label = QLabel("Rolling Average Window:")
        self.line_edit = QLineEdit()

        # Initial Horizontal layout used to order the labels and editor
        text_edit_layout = QHBoxLayout()
        text_edit_layout.addWidget(text_editor_label)
        text_edit_layout.addWidget(self.line_edit)

        # Adding the button to receive the data input by user
        set_time_button = QPushButton("Set Window")
        set_time_button.clicked.connect(self.set_window)

        # Vertical layout used to place button below text editor
        full_layout = QVBoxLayout()
        full_layout.addLayout(text_edit_layout)
        full_layout.addWidget(set_time_button)

        # Sets window layout
        self.setLayout(full_layout)

    def set_window(self):
        try:
            int(self.line_edit.text())
            self.impedance_smoothing.emit(self.line_edit.text())
            self.close()
        except ValueError:
            string_error = QMessageBox()
            string_error.setWindowTitle("Error")
            string_error.setText("Only Use Digits")
            string_error.setIcon(QMessageBox.Icon.Critical)
            string_error.setDefaultButton(QMessageBox.StandardButton.Ok)
            string_error.exec()


class HelpWidget(QPdfView):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Help Document")  # Set Window Title
        self.setWindowIcon(QIcon("Resources\\HelpIcon.png"))
        self.resize(850, 500)
        self.help_pdf = QPdfDocument()
        self.help_pdf.load("Resources\\GlucoseMeasuringHelp.pdf")  # Loads path of help document
        self.setPageMode(QPdfView.PageMode.MultiPage)
        self.setDocument(self.help_pdf)
