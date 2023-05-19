import os, random
import re
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtCore import Qt
import xlsxio
from datetime import datetime
import sys, traceback
import zipfile
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtSql import QSqlRelationalDelegate, QSqlQuery, QSqlRelationalTableModel, QSqlRelation, QSqlDatabase
from PyQt5.QtWidgets import QTableWidget, QGroupBox, QRadioButton, QLineEdit, QFileDialog, QProgressBar, QGridLayout, QMessageBox, QListWidget, QAbstractItemView, \
    QListWidgetItem, QLabel, QToolBar, QAction, QMainWindow, QWidget, QMenu, QMenuBar, QDialog, QPushButton, \
    QDialogButtonBox, QVBoxLayout, QHBoxLayout, QTextEdit, QTableView, QHeaderView, QComboBox, QSizePolicy
from pyexcel_xlsx import save_data
from collections import OrderedDict

startTime = datetime.now()  # subtotales
startTime2 = datetime.now()  # total acumulado

# variables globales
cadena = ""  # taskID
tiempoTotal = 0
lineasTotales = 0
lista_a_borrar = []
ruta = ""
nombreArchivo = ""
isSelected = False
isTabsOK = False
data = OrderedDict()
fileName = ""
pathSinFileName = ""
isExcelCreated = False
# fin variables globales

# Set up style sheet for the entire GUI
style_sheet = '''
    QWidget{
    }
    QMenuBar{
        background-color: #232a40;;
        font-size: 16px;
    }
    QLabel{
        background-color: pink;
        border-width: 2px;
        border-style: solid;
        border-radius: 8px;
        border-color: purple;
        color: black;
        font-size: 16px;
    }
    QTextEdit{
        background-color: pink;
        border-width: 2px;
        border-style: solid;
        border-radius: 8px;
        border-color: green;
        padding-left: 10px;
        color: black;  
        font-size: 16px;
    }
     QLineEdit{
        background-color: pink;
        border-width: 2px;
        border-style: solid;
        border-radius: 8px;
        border-color: green;
        padding-left: 10px;
        color: black;  
        font-size: 16px;
    }
    QListWidget{
        background-color: #a5a6b0;
        border-width: 2px;
        border-style: solid;
        border-radius: 8px;
        border-color: green;
        padding-left: 10px;
        color: #961A07;
    }    
    QTableView{
        background-color: #a5a6b0;
        border-width: 2px;
        border-style: solid;
        border-radius: 8px;
        border-color: green;
        padding-left: 10px;
        color: #961A07;
    }
    QPushButton{
        background-color: black;
        border-radius: 8px;
        padding: 6px;
        color: #FFFFFF;
        font-size: 18px;
    }
    QPushButton#initiate{
        background-color: black;
        border-radius: 8px;
        padding: 6px;
        color: #FFFFFF;
        font-size: 18px;
    }

    QPushButton:pressed{
        background-color: #C86354;
        border-radius: 4px;
        padding: 6px;
        color: #DFD8D7
    }
    QPushButton:hover {
        background-color: #0d1f54;
        border: 1px solid #0062cc;
    }
    QProgressBar{
        background-color: #C0C6CA;
        color: #FFFFFF;
        border: 1px solid grey;
        border-radius: 8px;
        padding: 3px;
        height: 15px;
        text-align: center;
    }

'''

class MainGUI(QtWidgets.QWidget):
    """
    This class holds all widgets and handles main interaction with user with a GUI
    """
    def __init__(self):
        """
        Setting up our GUI
        """
        super().__init__()
        self.initializeUI()


    def initializeUI(self):
        """
        Set up the application's GUI.
        """
        self.setFixedSize(640, 290)
        #self.setMinimumSize(640, 280)
        #self.setWindowTitle("AMM-Exploiter")
        self.setUpMainWindow()
        self.show()

    def setUpMainWindow(self):
        """
        Create and arrange widgets in the main window
        """
        # 3 text lines
        self.title = QtWidgets.QLabel("1.  File>Open Excel file")
        #self.title.setAlignment(QtCore.Qt.AlignCenter)
        self.title.setStyleSheet('font-size: 18px;')

        self.title2 = QtWidgets.QLabel("2.  Enter TASK-ID to search")
        #self.title2.setAlignment(QtCore.Qt.AlignCenter)
        self.title2.setStyleSheet('color: gray; font-size: 18px;')

        self.subtitle = QtWidgets.QLabel("Selected file:\nNone")
        self.subtitle.setWordWrap(True)
        #self.subtitle.setAlignment(QtCore.Qt.AlignCenter)
        self.subtitle.setStyleSheet('color: #3a9c3b; font-size: 16px;')

        self.listtitle = QtWidgets.QLabel("Tasks Queue:")
        self.listtitle.setAlignment(QtCore.Qt.AlignCenter)
        self.listtitle.setStyleSheet('font-size: 16px;')

        # 1 input box
        self.input_box = QtWidgets.QLineEdit()
        if (isSelected == False):
            self.input_box.setEnabled(False)
        self.input_box.setPlaceholderText("00-00-00-000-000-A")

        # 3 buttons
        self.button = QtWidgets.QPushButton("Initiate")
        self.button.setObjectName("initiate")
        self.button.setToolTip('Start processing')
        self.button.setStyleSheet('''
                                           QPushButton:hover {
                                                background-color: #0d1f54;
                                                border: 1px solid #0062cc;
                                           }
                                       ''')
        self.button.clicked.connect(self.start_task)

        #
        self.button2 = QtWidgets.QPushButton("Select file")
        self.button2.setToolTip('Select Excel file to process')
        self.button2.clicked.connect(self.abrirArchivo)
        self.button2.setStyleSheet('''
                                           QPushButton:hover {
                                               background-color: #3a9c3b;
                                               border: 1px solid #0062cc;
                                           }
                                       ''')

        #
        self.boton_generar = QtWidgets.QPushButton("Open generated file")
        self.boton_generar.setToolTip('Open up Excel generated file')
        self.boton_generar.clicked.connect(self.abrirGenerado)
        self.boton_generar.setStyleSheet('''
                                           QPushButton:hover {
                                               background-color: #a30393;
                                               border: 1px solid #0062cc;
                                           }
                                       ''')

        #
        self.btn_addTask = QtWidgets.QPushButton(">>")
        self.btn_addTask.setToolTip("Add task-Id to queue for multi-tasks processing")
        self.btn_addTask.clicked.connect(self.addTaskToList)
        self.btn_addTask.setStyleSheet('''
                                           QPushButton:hover {
                                               background-color: #3a9c3b;
                                               border: 1px solid #0062cc;
                                           }
                                       ''')

        #
        self.btn_removeTask = QtWidgets.QPushButton("<<")
        self.btn_removeTask.setToolTip("Remove selected task-Id from queue")
        self.btn_removeTask.clicked.connect(self.removeTaskFromList)
        self.btn_removeTask.setStyleSheet('''
                                           QPushButton:hover {
                                               background-color: #ed0c0c; 
                                               border: 1px solid #0062cc;
                                           }
                                       ''')

        # creating a QListWidget
        self.list_widget = QListWidget(self)
        self.list_widget.setObjectName("listw")
        self.list_widget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.list_widget.itemSelectionChanged.connect(self.on_change)

        # setting geometry to list
        #self.list_widget.setGeometry(50, 70, 250, 200)
        # list widget items
        item1 = QListWidgetItem("00-00-00-000-000-A")

        # 1 progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)


        self.business_logic = BusinessLogic()
        self.worker_thread = WorkerThread(self.business_logic)
        self.worker_thread.progress_changed.connect(self.update_progress_bar)
        self.worker_thread.task_completed.connect(self.show_message_box)

        # self.setStyleSheet('''
        #
        #                     QWidget {
        #                         background-color: #F5F5F5;
        #                     }
        #                     QLineEdit {
        #                         background-color: white;
        #                         border: 1px solid #ccc;
        #                         border-radius: 5px;
        #                         padding: 5px;
        #                         font-size: 18px;
        #                     }
        #
        #                     QPushButton {
        #                         background-color:  #000000;
        #                         border: 1px solid #007bff;
        #                         border-radius: 5px;
        #                         color: white;
        #                         font-size: 18px;
        #                         padding: 5px 10px;
        #                     }
        #                     QPushButton:hover {
        #                         background-color: #0069d9;
        #                         border: 1px solid #0062cc;
        #                     }
        #                 ''')
        # self.setLayout(layout)

        # Create layout and arrange widgets
        grid = QGridLayout()
        grid.addWidget(self.title, 0, 0,1,2)
        grid.addWidget(self.listtitle, 0, 3)
        grid.addWidget(self.title2, 1, 0)
        grid.addWidget(self.subtitle, 2, 0)
        grid.addWidget(self.input_box, 3, 0, 1, 2)
        grid.addWidget(self.btn_removeTask, 2, 2, 1, 1)
        grid.addWidget(self.btn_addTask, 3, 2, 1, 1)
        grid.addWidget(self.list_widget, 1, 3, 3, 1)
        grid.addWidget(self.button, 4, 0, 1, 4)
        grid.addWidget(self.progress_bar, 5, 0, 1, 4)

        self.setLayout(grid)

    def on_change(self):
        print("start")
        global lista_a_borrar
        lista_a_borrar.clear()
        #print([item.text() for item in self.list_widget.selectedItems()])
        for item in self.list_widget.selectedItems():
            lista_a_borrar.append(item.text())
        print(lista_a_borrar)

    def addTaskToList(self):
        """
        Add task-id from input box to list, after processing its format
        :return:
        """
        if self.business_logic.isValid1(self.input_box.text()): # falta cambiar para permitir initiate
            print("Task-ID format is ok")
            if self.input_box.text() != "":  # aqui verificar que cumple formato
                self.list_widget.addItem(self.input_box.text())
                self.input_box.setText("")
        else:
            print("Task-ID format is NOT ok")
            self.business_logic.mostrar_error("Error", "Task-ID format is not correct")

    def removeTaskFromList(self):
        """
        Remove selected task-id from list
        :return:
        """
        for x in self.list_widget.selectedIndexes():
            print(f"borramos: {x.row()}")
            self.list_widget.takeItem(x.row())


    def start_task(self):
        """
        Triggers actual thread execution in run function
        :return:
        """

        global cadena
        if self.input_box.text() != "":
            cadena = self.input_box.text()

        self.progress_bar.setValue(0)
        isOk = self.procesaExcel()
        if isOk:
            self.worker_thread.start()

    def update_progress_bar(self, value):
        """
        Updates progressBar value in GUI
        :param value:
        :return:
        """
        self.progress_bar.setValue(value)

    def show_message_box(self):
        """
        Pops up a final message when processing is finished without errors
        :return:
        """
        QMessageBox.information(self, "Successfully done", "Completed!")

    def procesaExcel(self):
        """
        Caller function
        :return:
        """
        # cadena = "05-21-00-200-802-A"
        global cadena
        if self.input_box.text() != "":
            cadena = self.input_box.text()  # asignar listwisget[0] si input es ""
        else:
            lst = self.list_widget
            items = []
            for x in range(lst.count()):
                items.append(lst.item(x).text())
                cadena = lst.item(x).text()
            print(items)
            # revisar esto, como acceder

        global isExcelCreated
        isExcelCreated = False
        global isFormatOK
        isFormatOK = False
        if self.business_logic.isValid1(cadena) or self.list_widget.count() > 0: # coger la lista de widget
            print("Task-ID format is ok")
            isFormatOK = True
            return True

        else:
            print("Task-ID format is NOT ok")
            self.business_logic.mostrar_error("Error", "Task-ID format is not correct")
            return False

    def abrirGenerado(self):
        """
        Opens up the generated Excel file to be browsed by user
        :return:
        """
        global isExcelCreated
        if (not isExcelCreated):
            print(f"No file has been generated yet, status \"isExcelCreated\":{isExcelCreated}")
        if (os.path.exists(pathSinFileName + cadena + ".xlsx")) and isExcelCreated:  # poner filepath
            os.startfile(pathSinFileName + cadena + ".xlsx")

    def abrirArchivo(self):
        """
        Load a local file by reference to perform a later processing. Selected file is shown in GUI
        :return:
        """
        archivo, _ = QFileDialog.getOpenFileName(self, 'Select Excel file (.xlsx) to process', filter='Excel files (*.xlsx)')

        if archivo:
            global isSelected
            isSelected = True
            global ruta
            ruta = archivo
            global nombreArchivo
            nombreArchivo = ruta.split('/')[-1]
            print(f'Selected file {nombreArchivo}\n located in: {ruta}')

            if isSelected == True:
                self.input_box.setEnabled(True)
                self.subtitle.setText("Selected file:\n" + nombreArchivo)
                self.title.setStyleSheet('color: gray; font-size: 20px;')
                self.title2.setStyleSheet('color: black; font-size: 20px;')
                #QtTest.QTest.qWait(1000)

            global isTabsOK
            isTabsOK = False
            self.business_logic.comprobarTabs(ruta) # checks if selected file has required tabs'
            # Here starts autocomplete as follows: if istabs ok, process Tools tab, save all tasksID in a global set
            # and block the editbox by showing a display msg saying waiting to autocomplete. Later, it leaves
            # editbox like before, available to type, and with hint example. Incluir menu para elegir batch o single
            # task-ID search.

class BusinessLogic:
    """
    This class encloses almost all of the processing functions of our domain
    """
    def comprobarTabs(self, ruta):
        """
                Checks wether provided Excel file complies with required tabs structure
                :param ruta: (string) file path
                :return:
                """
        print(f"Check the file: {ruta}")
        sheets = self.getTabNames(ruta)

        if "Tools" and "Consumables" and "Expendables" and "IPC" and "TASK" in sheets:

            print(
                "Ok, Excel file complies with expected tabs' names:\n['Tools', 'Consumables', 'Expendables', 'IPC', 'TASK']")
            global isTabsOK
            isTabsOK = True

        else:

            print(
                "Error: Excel file does NOT comply with expected tabs' names\n['Tools', 'Consumables', 'Expendables', 'IPC', 'TASK']")
            self.mostrar_error("Error", "Selected file must contain following tabs' names: \n\nTools, Consumables, Expendables, IPC and TASK")
            # aquí no continuar - mostrar aviso

    def getTabNames(self, file_path):
        """
        Gets actual tabs' names of provided excel file
        :param file_path: (string) file path
        :return:
        """
        sheets = []
        with zipfile.ZipFile(file_path, 'r') as zip_ref: xml = zip_ref.read("xl/workbook.xml").decode("utf-8")
        for s_tag in re.findall("<sheet [^>]*", xml): sheets.append(re.search('name="[^"]*', s_tag).group(0)[6:])
        return sheets

    def crearExcel(self):
        """
        Generates output Excel file once processing is done successfully
        :return:
        """
        global fileName
        fileName = ruta.split("/")[-1]
        print(f"{fileName}")
        global pathSinFileName
        pathSinFileName = ruta.replace(fileName, "")
        print(f"{pathSinFileName}")

        try:
            global data
            global cadena  # to concat date and hour
            data.move_to_end("IPC")  # fixed to show correct order
            data.move_to_end("TASK")
            save_data(pathSinFileName + cadena + ".xlsx", data)
            keys = list(data.keys())
            print(keys)
            global isExcelCreated
            isExcelCreated = True
            print(f"Excel file has been generated: {cadena}.xlsx in: {pathSinFileName}")
            self.thread.do_progress(100)
        except PermissionError as perr:
            print("File is open, please close it", perr)
            self.mostrar_error("Error",
                               "File is currently opened, must be closed in order to be generated")

    def buscarEnExcel(self, taskId):
        """
        Caller function that calls all main processing functions
        :param taskId:
        :return:
        """
        global isTabsOK
        if isTabsOK:
            global startTime2
            startTime2 = datetime.now()
            global lineasTotales
            lineasTotales = 0
            global data
            data.clear()  # delete previous when multiple searches are done while app opened
            self.explorarTools(taskId)
            self.explorarConsumables(taskId)
            self.explorarExpendables(taskId)
            self.explorarTASK(taskId)
            self.imprimirTiempos()
            self.crearExcel()

    def explorarTools(self, taskId):
        """
        Processes "Tools" tab from provided file
        :param taskId: (string)
        :return:
        """
        print("*********************************************")
        print(f"Searching that task-ID in Tools: {taskId}")
        print("*********************************************")
        startTime = datetime.now()
        total_filas = 0
        veces_encontrado = 0
        encontrado = False
        pos = 1  # empiezan los datos en la 1
        celdaText = ""
        celdaQty = ""
        celdaDesignation = ""
        self.thread.do_progress(10)
        listaFinal = []
        types = [str, str, str, str, str, str, str, str, str, str, str, str, str]  # col number
        try:
            with xlsxio.XlsxioReader(ruta) as reader:
                with reader.get_sheet('Tools', types=types) as sheet:
                    header = sheet.read_header()
                    print(header)
                    only_active = []
                    for row in sheet.iter_rows():
                        total_filas += 1
                        pos += 1
                        if row[5] == taskId:
                            # only_active.append(row)
                            # print(only_active)
                            celdaText = row[8]
                            if (9 < len(row)):
                                celdaQty = row[9]
                            if (10 < len(row)):  # safe access
                                celdaDesignation = row[10]
                            veces_encontrado += 1
                            subLista = [celdaText, celdaQty, celdaDesignation]
                            listaFinal.append(subLista)
                            print(
                                f"Encontrado en fila: {pos}, Text={celdaText}, Qty:{celdaQty}, Design:{celdaDesignation}")
                            encontrado = True
        except:
            print("************ START ERROR ***************")
            traceback.print_exc(file=sys.stdout)
            print("************ END ERROR ******************")
        global data
        data.update({"Tools": listaFinal})
        print("Time, subtotal: {}".format((datetime.now() - startTime)))
        global lineasTotales
        lineasTotales += total_filas
        print(f"Total file rows: {total_filas}")
        if (encontrado == False):
            print("Not found")
        self.thread.do_progress(20)

    def explorarConsumables(self, taskId):
        """
        Processes "Consumables" tab from provided file
        :param taskId:
        :return:
        """
        print("*********************************************")
        print(f"Searching task-ID in Consumables: {taskId}")
        print("*********************************************")
        startTime = datetime.now()
        total_filas = 0
        veces_encontrado = 0
        encontrado = False
        pos = 1  # data start at position 1
        celdaConnbr = ""
        celdaName = ""
        listaFinal = []
        types = [str, str, str, str, str, str, str, str, str, str, str, str, str]  # 13, col number
        try:
            with xlsxio.XlsxioReader(ruta) as reader:
                with reader.get_sheet('Consumables', types=types) as sheet:
                    header = sheet.read_header()
                    print(header)

                    only_active = []
                    for row in sheet.iter_rows():

                        total_filas += 1
                        pos += 1
                        if row[4] == taskId:
                            celdaConnbr = row[9]
                            celdaName = row[11]
                            veces_encontrado += 1

                            subLista = [celdaConnbr, celdaName]
                            listaFinal.append(subLista)
                            #
                            print(f"Encontrado en fila: {pos}, Connbr={celdaConnbr}, Name={celdaName}")
                            encontrado = True

            print("Tiempo subtotal: {}".format((datetime.now() - startTime)))
            print(f"Total filas del archivo: {total_filas}")

            global lineasTotales
            lineasTotales += total_filas

            if (encontrado == False):
                print("Not found")

        except:
            print("************ START ERROR ***************")
            traceback.print_exc(file=sys.stdout)
            print("************ END ERROR ******************")

        global data
        data.update({"Consumables": listaFinal})
        self.thread.do_progress(40)

    def explorarIPC(self, cadenaFirst8, cola3):
        """
        Processes "IPC" tab from provided file
        :param cadenaFirst8:
        :param cola3:
        :return:
        """
        # - ITEM de solo 2 cifras o con 4. FIG con letra al final
        print(f"String: {cadenaFirst8} and tail of 3 chars: {cola3}")
        print("****************************************************")
        print(f"Searching FIG_REFERENCE {cadenaFirst8} in IPC tab")
        print("****************************************************")

        startTime = datetime.now()
        total_filas = 0
        veces_encontrado = 0
        encontrado = False
        pos = 1  # data start at position 1
        celdaPart_number = ""
        celdaUnit_per_assy = ""
        celdaSpare = ""
        celdaItem = ""
        listaFinal = []
        self.thread.do_progress(70)
        types = [str, str, str, str, str, str, str, str, str, str, str, str]  # col number
        try:
            with xlsxio.XlsxioReader(ruta) as reader:
                with reader.get_sheet('IPC', types=types) as sheet:
                    header = sheet.read_header()
                    print(header)
                    only_active = []

                    for row in sheet.iter_rows():
                        total_filas += 1
                        pos += 1

                        if (row[4][:11] == cadenaFirst8):  # and row[5] == cola3:
                            celdaItem = row[5]
                            celdaItem = celdaItem.zfill(3)  # zeros filling

                            if (celdaItem.endswith(".0")):  # fix .0
                                celdaItem = celdaItem[:celdaItem.index('.')]  # remove .0
                                celdaItem = celdaItem.zfill(3)

                            if (celdaItem[:3] == cola3):  # match, FIG and Item
                                encontrado = True
                                veces_encontrado += 1
                                celdaPart_number = row[6]
                                celdaUnit_per_assy = row[7]
                                celdaSpare = row[8]
                                global data
                                if (celdaSpare == "#########"):
                                    # saving data
                                    subLista = [celdaPart_number, celdaUnit_per_assy, "-"]
                                    listaFinal.append(subLista)
                                    print(
                                        f"Found in row: {pos}, FIG:{row[4]}, ITEM:{celdaItem}, PN={celdaPart_number}, Unit:{celdaUnit_per_assy}, Spare: - ")
                                else:
                                    subLista = [celdaPart_number, celdaUnit_per_assy, celdaSpare]
                                    listaFinal.append(subLista)
                                    print(
                                        f"Found in row: {pos}, FIG:{row[4]}, ITEM:{celdaItem}, PN={celdaPart_number}, Unit:{celdaUnit_per_assy}, Spare:{celdaSpare} ")

        except:
            print("************ START ERROR ***************")
            traceback.print_exc(file=sys.stdout)
            print("************ END ERROR ******************")

        global data
        data.update({"IPC": listaFinal})
        print("Time, subtotal: {}".format((datetime.now() - startTime)))
        print(f"Total file rows: {total_filas}")
        global lineasTotales
        lineasTotales += total_filas
        print(f"This: {veces_encontrado} ocurrences found")
        if (encontrado == False):
            print("Not found")
        self.thread.do_progress(80)

    def procCSN(self, celdaCSN):
        """
        Processes "CSN" field in tab IPC from provided file
        :param celdaCSN:
        :return:
        """
        cadenaFirst8 = ""
        cola3 = ""
        if (self.isValid2(celdaCSN)):
            cadenaFirst8 = self.insertaGuiones(celdaCSN)
            cola3 = self.tail3(celdaCSN)
            self.explorarIPC(cadenaFirst8, cola3)

    def explorarExpendables(self, taskId):
        """
        Processes "Expendables" tab from provided file
        :param taskId: (string)
        :return:
        """
        print("*****************************************************")
        print(f"Searching task-id: {taskId} in Expendables tab")
        print("*****************************************************")

        startTime = datetime.now()
        total_filas = 0
        veces_encontrado = 0
        encontrado = False
        pos = 1  # empiezan los datos en la 1
        celdaCSN = ""
        celdaCSNProcesada = ""
        celdaName = ""
        listaFinal = []

        types = [str, str, str, str, str, str, str, str, str, str, str, str, str]  # 13, cols number
        try:
            with xlsxio.XlsxioReader(ruta) as reader:
                with reader.get_sheet('Expendables', types=types) as sheet:
                    header = sheet.read_header()
                    print(header)
                    only_active = []
                    for row in sheet.iter_rows():
                        total_filas += 1
                        pos += 1
                        if row[4] == taskId:
                            # only_active.append(row)
                            # print(only_active)
                            celdaCSN = row[9]
                            print(f"Celda CSN: {celdaCSN}")
                            self.procCSN(celdaCSN)
                            celdaName = row[11]
                            veces_encontrado += 1
                            subLista = [celdaName]
                            listaFinal.append(subLista)
                            print(f"Found in row: {pos}, CSN={celdaCSN}, Name={celdaName}")
                            encontrado = True
            print("Time, subtotal: {}".format((datetime.now() - startTime)))
            print(f"Total file rows: {total_filas}")
            global lineasTotales
            lineasTotales += total_filas

            if (encontrado == False):
                subLista = ['']
                listaFinal.append(subLista)  # if not found, creates empty one
                print("Not found")

        except:
            print("************ START ERROR ***************")
            traceback.print_exc(file=sys.stdout)
            print("************ END ERROR ******************")
        global data
        data.update({"Expendables": listaFinal})
        if encontrado == False:
            data.update({"IPC": listaFinal})

    def explorarTASK(self, taskId):
        """
        Processes "TASK" tab from provided file
        :param taskId:
        :return:
        """
        print("************************************************")
        print(f"Searching task-ID: {taskId} in TASK tab")
        print("************************************************")
        startTime = datetime.now()
        total_filas = 0
        veces_encontrado = 0
        encontrado = False
        pos = 1  # empiezan los datos en la 1
        celdaRefid = ""
        celdaSubtask = ""
        listaFinal = []

        types = [str, str, str, str, str, str, str, str, str, str, str, str, str]  # cols number
        try:
            with xlsxio.XlsxioReader(ruta) as reader:
                with reader.get_sheet('TASK', types=types) as sheet:
                    header = sheet.read_header()
                    print(header)
                    only_active = []
                    for row in sheet.iter_rows():
                        total_filas += 1
                        pos += 1
                        if row[4] == taskId:
                            # only_active.append(row)
                            # print(only_active)
                            celdaRefid = row[9]
                            celdaSubtask = row[10]
                            veces_encontrado += 1
                            subLista = [celdaRefid, celdaSubtask]
                            listaFinal.append(subLista)
                            print(
                                f"Encontrado en fila: {pos}, Refid={celdaRefid}, SubTask:{celdaSubtask}")
                            encontrado = True
        except:
            print("************ START ERROR ***************")
            traceback.print_exc(file=sys.stdout)
            print("************ END ERROR ******************")

        global data
        data.update({"TASK": listaFinal})
        print("Tiempo subtotal: {}".format((datetime.now() - startTime)))
        global lineasTotales
        lineasTotales += total_filas
        print(f"Total filas del archivo: {total_filas}")
        if (encontrado == False):
            print("No encontrado")
            subLista = ['']
            listaFinal.append(subLista)  #
            data.update({"TASK": listaFinal})
        self.thread.do_progress(99)

    def mostrar_error(self, titulo, textoerror):
        """
        Pops up error messages
        :param titulo:
        :param textoerror:
        :return:
        """
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText(textoerror)
        msg.setWindowTitle(titulo)
        msg.exec_()

    def run(self, taskId):
        """
        Thread actual execution
        :param taskId: (str)
        :return:
        """
        # time counter starts here
        global isTabsOK
        global isFormatOK
        if (isTabsOK == True) and (isFormatOK == True):
            global startTime2
            startTime2 = datetime.now()
            global lineasTotales
            lineasTotales = 0
            global data
            data.clear()  # delete previous if multiple searchings
            self.explorarTools(taskId)
            self.explorarConsumables(taskId)
            self.explorarExpendables(taskId)
            self.explorarTASK(taskId)
            self.imprimirTiempos()
            self.crearExcel()

    def insertaGuiones(self, cadena):
        """
            Insert hyphens in the string "\d{8}-\d{3}"
            :param cadena: (str) input string
            :return:
                str: string like: 12-34-56-789-123
            """
        nuevaCadena = cadena[:2] + "-" + cadena[2:4] + "-" + cadena[4:6] + "-" + cadena[6:8]
        return nuevaCadena

    def isValid2(self, cadena):
        """
        Checks wether provided string has proper format: 123456789-123
        :param cadena: (str) Cadena de entrada
        :return:
            Bool: La cadena cumple el formato "\d{8}-\d{3}"
        """
        regex = re.search("\d{8}-\d{3}", cadena)
        return regex

    def tail3(self, cadena):
        """
        Crop last 3 chars
        :param cadena:
        :return:
            str: last 3 chars of input string
        """
        return cadena[9:12]

    def isValid1(self, cadena):
        """
        Checks if provided string complies with format
        :param cadena: (str)
        :return:
        """
        regex = re.search("\d{2}-\d{2}-\d{2}-\d{3}-\d{3}-[A-Z]", cadena)
        return regex

    def imprimirTiempos(self):
        """
        Prints out elapsed time in the console
        :return:
        """
        print("*****************************************")
        print("TOTAL TIME: {}".format((datetime.now() - startTime2)))
        print(f"PROCESSED ROWS:{lineasTotales}")
        print("*****************************************")

class WorkerThread(QThread):
    """
    Thread class
    """
    progress_changed = pyqtSignal(int)
    task_completed = pyqtSignal()

    def __init__(self, business_logic):
        super().__init__()
        self.business_logic = business_logic

    def run(self):
        """
        In this function is where actual thread starts off
        :return:
        """
        self.business_logic.thread = self
        self.business_logic.run(cadena)
        self.task_completed.emit()

    def do_progress(self, value):
        """
        Update progress bar value
        :param value: (int) value for progress bar
        :return:
        """
        self.progress_changed.emit(value)

class MainWindow(QMainWindow):
    """
    This class contains the GUI design from classes which inherits from QWdidget, to be able to add a Menu
    """

    def __init__(self):
        super().__init__()
        self.initializeUI()

    def initializeUI(self):
        """Set up the application's GUI."""
        self.mainForm = MainGUI()
        self.setCentralWidget(self.mainForm)
        self.setWindowTitle("AMM-Exploiter")
        self.setFixedSize(640, 310)

        # Crear menú
        menuFile = self.menuBar()
        menuFile.setNativeMenuBar(False) # just for macOS, comment it out for Windows
        menuFile = self.menuBar().addMenu('File')
        menuFile.setStyleSheet('font-size: 16px;')

        subOp = QAction('Open Excel File', self)
        subOp.triggered.connect(self.mainForm.abrirArchivo)
        menuFile.addAction(subOp)

        subSettings = QAction('Settings', self)
        subSettings.triggered.connect(self.launchSettings)
        menuFile.addAction(subSettings)

        menuTools = self.menuBar().addMenu('Tools')
        menuTools.setObjectName("menutools")
        menuTools.setStyleSheet('font-size: 16px;')

        # Agregar subítem "Open"
        action = QAction('Add New Note', self)
        action.triggered.connect(self.addNewNote)
        menuTools.addAction(action)

        # Agregar subítem "Working Directory"
        action_imprimir = QAction('Watch All Notes', self)
        action_imprimir.triggered.connect(self.openAllNotes)
        menuTools.addAction(action_imprimir)


        self.show()

    def launchSettings(self):
        self.settings = Settings()
        if self.settings.exec():
            print("settings saved")
        else:
            print("Canceled!")


    def addNewNote(self):
        # not dialog add new note
        self.newNote = NewNote()
        self.newNote.show()

    def openAllNotes(self):
        # not dialog watch all notes
        self.watchNotes = WatchNotes()
        self.watchNotes.show()

    def openFileDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self, "Open File", "",
                                                  "All Files (*);;Python Files (*.py)", options=options)
        if fileName:
            print(f'Se seleccionó el archivo: {fileName}')


    def removeListWidget(self):
        for x in self.mainForm.list_widget.selectedIndexes():
            print(f"borramos: {x.row()}")
            self.mainForm.list_widget.takeItem(x.row())


    def addListWidget(self):
        self.mainForm.list_widget.addItem("Añadido")


class Settings(QDialog):
    """
    This class holds all widgets in Settings Window
    """
    def __init__(self):
        """
        Setting up GUI for settings window, it will be launched when clicking File>Settings
        """
        super().__init__()
        self.initializeUI()

    def initializeUI(self):
        """
        Set up the application's Settins GUI.
        """
        self.setWindowTitle("Settings")
        self.setMinimumSize(450, 200)
        self.btn = QDialogButtonBox.Cancel | QDialogButtonBox.Save
        self.buttonBox = QDialogButtonBox(self.btn)
        self.buttonBox.accepted.connect(self.saveSettings)
        self.buttonBox.rejected.connect(self.cancelSettings)

        # Create group box to contain radio buttons
        dir_gb = QGroupBox("Working directory:")

        self.inputDir = QLabel("C:/...")
        self.inputDir.setWordWrap(True)
        btn_chDir = QPushButton("Change")
        btn_chDir.clicked.connect(self.changeDir)
        line_dir = QHBoxLayout()
        line_dir.addWidget(self.inputDir, 3)
        line_dir.addWidget(btn_chDir, 1)
        dir_gb.setLayout(line_dir)


        # Create group box to contain radio buttons
        files_gb = QGroupBox("Generated Excel file(s):")

        self.single_rb = QRadioButton("Single")
        self.multi_rb = QRadioButton("Multi")
        self.multi_rb.setChecked(True)

        # Create and set layout for sex_gb widget
        files_h_box = QHBoxLayout()
        files_h_box.addWidget(self.single_rb)
        files_h_box.addWidget(self.multi_rb)

        files_gb.setLayout(files_h_box)

        self.layout = QVBoxLayout()
        self.layout.addWidget(dir_gb)
        self.layout.addWidget(files_gb)
        self.layout.addWidget(self.buttonBox)
        self.setLayout(self.layout)

    def changeDir(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        directory = QFileDialog.getExistingDirectory(self, "Select Directory", options=options)
        if directory:
            self.inputDir.setText(directory)

    def saveSettings(self):
        print("saving settings")
        if self.single_rb.isChecked():
            mode = "Single"
        else:
            mode = "Multi"
            # Aquí puedes guardar la configuración en función del modo seleccionado
        print("Modo seleccionado:", mode)
        self.accept()

    def cancelSettings(self):
        print("canceling settings")
        self.close()

class WatchNotes(QWidget):
    def __init__(self):
        super().__init__()
        # CreateEmployeeData()
        self.initializeUI()

    def initializeUI(self):
        """
        Initialize the window and display its contents to the screen.
        """
        self.setMinimumSize(800, 600)
        self.setWindowTitle("Watch All Comments")
        #self.createTables()
        self.createConnection()
        self.createTable()
        self.setupWidgets()

        self.show()


    def createConnection(self):
        if QSqlDatabase.contains("qt_sql_default_connection"):
            return True  # Ya hay una conexión activa, no es necesario volver a conectar
        database = QSqlDatabase.addDatabase("QSQLITE")  # SQLite version 3
        database.setDatabaseName("files/accounts.db")
        #database.setDatabaseName(r"\\gfa60005\intercambios\AlRa\db\accounts.db")

        if not database.open():
            print("-Unable to open data source file.")  # this happens when the directory itself does not exist
            sys.exit(1)  # Error code 1 - signifies error

        # Check if the tables we need exist in the database
        tables_needed = {'accounts', 'countries'}
        tables_not_found = tables_needed - set(database.tables())
        if tables_not_found:
            QMessageBox.critical(None, 'Error',
                                 f'The following tables tables are missing from the database: {tables_not_found}')
            sys.exit(1)  # Error code 1 - signifies error

    def createTable(self):
        """
        Set up the model, headers and populate the model.
        """
        self.model = QSqlRelationalTableModel()
        self.model.setTable('accounts')
        self.model.setRelation(self.model.fieldIndex('country_id'), QSqlRelation('countries', 'id', 'country'))

        self.model.setHeaderData(self.model.fieldIndex('id'), Qt.Horizontal, "ID")
        self.model.setHeaderData(self.model.fieldIndex('employee_id'), Qt.Horizontal, "Employee ID")
        self.model.setHeaderData(self.model.fieldIndex('first_name'), Qt.Horizontal, "First")
        self.model.setHeaderData(self.model.fieldIndex('last_name'), Qt.Horizontal, "Last")
        # self.model.setHeaderData(self.model.fieldIndex('email'), Qt.Horizontal, "E-mail")
        self.model.setHeaderData(self.model.fieldIndex('department'), Qt.Horizontal, "Dept.")
        self.model.setHeaderData(self.model.fieldIndex('country_id'), Qt.Horizontal, "Country")

        # Populate the model with data
        self.model.select()

    def setupWidgets(self):
        """
        Create instances of widgets, the table view and set layouts.
        """
        icons_path = "icons"

        title = QLabel("List of all comments")
        title.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        title.setStyleSheet("font: bold 20px")

        add_product_button = QPushButton("Add Employee")
        add_product_button.setIcon(QIcon(os.path.join(icons_path, "add_user.png")))
        add_product_button.setStyleSheet("padding: 10px")
        add_product_button.clicked.connect(self.addItem)

        del_product_button = QPushButton("Delete")
        del_product_button.setIcon(QIcon(os.path.join(icons_path, "trash_can.png")))
        del_product_button.setStyleSheet("padding: 10px")
        del_product_button.clicked.connect(self.deleteItem)

        # Set up sorting combobox
        sorting_options = ["Sort by ID", "Sort by Employee ID", "Sort by First Name",
                           "Sort by Last Name", "Sort by Department", "Sort by Country"]
        sort_name_cb = QComboBox()
        sort_name_cb.addItems(sorting_options)
        sort_name_cb.currentTextChanged.connect(self.setSortingOrder)

        buttons_h_box = QHBoxLayout()
        buttons_h_box.addWidget(add_product_button)
        buttons_h_box.addWidget(del_product_button)
        buttons_h_box.addStretch()
        buttons_h_box.addWidget(sort_name_cb)

        # Widget to contain editing buttons
        edit_buttons = QWidget()
        edit_buttons.setLayout(buttons_h_box)

        # Create table view and set model
        self.table_view = QTableView()
        self.table_view.setModel(self.model)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_view.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_view.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Deshabilitar la edición
        self.table_view.setSelectionMode(QTableView.SingleSelection)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)

        # Instantiate the delegate
        delegate = QSqlRelationalDelegate(self.table_view)
        self.table_view.setItemDelegate(delegate)

        # Main layout
        main_h = QHBoxLayout()

        main_v_box = QVBoxLayout()
        main_v_box.addWidget(title, Qt.AlignLeft)
        main_v_box.addWidget(edit_buttons)
        main_v_box.addWidget(self.table_view)

        main_v_box_der = QVBoxLayout()
        self.comments = QTextEdit()
        main_v_box_der.addWidget(self.comments)

        main_h.addLayout(main_v_box)
        main_h.addLayout(main_v_box_der)

        #self.table_view.itemSelectionChanged.connect(self.show_selected_comment)
        self.table_view.selectionModel().selectionChanged.connect(self.show_selected_comment)

        self.setLayout(main_h)

    def show_selected_comment(self):
        #selected_items = self.table_view.selectedItems()
        selected_indexes = self.table_view.selectionModel().selectedIndexes()
        selected_items = [index.data() for index in selected_indexes]
        print(selected_items)
        print(selected_indexes[0].sibling(selected_indexes[0].row(), 2).data())
        if selected_indexes:
            comment = selected_indexes[0].sibling(selected_indexes[0].row(), 2).data()
            self.comments.setText(comment)
        else:
            pass


    def addItem(self):
        """
        Add a new record to the last row of the table.
        """
        last_row = self.model.rowCount()
        self.model.insertRow(last_row)

        id = 0
        query = QSqlQuery()
        query.exec_("SELECT MAX (id) FROM accounts")
        if query.next():
            print(query.value(0))
            id = int(query.value(0))

    def deleteItem(self):
        """
        Delete an entire row from the table.
        """
        current_item = self.table_view.selectedIndexes()
        for index in current_item:
            self.model.removeRow(index.row())
        self.model.select()

    def setSortingOrder(self, text):
        """
        Sort the rows in table.
        """
        if text == "Sort by ID":
            self.model.setSort(self.model.fieldIndex('id'), Qt.AscendingOrder)
        elif text == "Sort by Employee ID":
            self.model.setSort(self.model.fieldIndex('employee_id'), Qt.AscendingOrder)
        elif text == "Sort by First Name":
            self.model.setSort(self.model.fieldIndex('first_name'), Qt.AscendingOrder)
        elif text == "Sort by Last Name":
            self.model.setSort(self.model.fieldIndex('last_name'), Qt.AscendingOrder)
        elif text == "Sort by Department":
            self.model.setSort(self.model.fieldIndex('department'), Qt.AscendingOrder)
        elif text == "Sort by Country":
            self.model.setSort(self.model.fieldIndex('country'), Qt.AscendingOrder)

        self.model.select()

class NewNote(QWidget):
    """
    This class holds all widgets in New Notes Window
    """

    def __init__(self):
        super().__init__()
        # CreateEmployeeData()
        self.initializeUI()

    def initializeUI(self):
        """
        Initialize the window and display its contents to the screen.
        """
        self.setMinimumSize(800, 600)
        self.setWindowTitle("Add new Comment")
        #self.createTables()
        self.createConnection()
        self.createTable()
        self.setupWidgets()

        self.show()

    def createTables(self):
        """
           Create sample database for project.
           Class demonstrates how to connect to a database, create queries,
           and create tables and records in those tables.
           """
        # Create connection to database. If db file does not exist,
        # a new db file will be created.
        database = QSqlDatabase.addDatabase("QSQLITE")  # SQLite version 3
        database.setDatabaseName("files/accounts.db")
        #database.setDatabaseName(r"\\gfa60005\intercambios\AlRa\db\accounts.db")

        if not database.open():
            print("Unable to open data source file.")
            sys.exit(1)  # Error code 1 - signifies error

        query = QSqlQuery()
        # Erase database contents
        query.exec_("DROP TABLE accounts")
        query.exec_("DROP TABLE countries")

        query.exec_("""CREATE TABLE accounts (
                             id INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL,
                             employee_id INTEGER NOT NULL,
                             first_name VARCHAR(30) NOT NULL,
                             last_name VARCHAR(30) NOT NULL,

                             department VARCHAR(20) NOT NULL, 
                             country_id VARCHAR(20) REFERENCES countries(id))""")

        # Positional binding to insert records into the database
        query.prepare("""INSERT INTO accounts (
                               employee_id, first_name, last_name, 
                               department, country_id) 
                               VALUES (?, ?, ?, ?, ?)""")

        first_names = ["Emma", "Olivia", "Ava", "Isabella", "Sophia",
                       "Mia", "Charlotte", "Amelia", "Evelyn", "Abigail",
                       "Valorie", "Teesha", "Jazzmin", "Liam", "Noah",
                       "William", "James", "Logan", "Benjamin", "Mason",
                       "Elijah", "Oliver", "Jason", "Lucas", "Michael"]

        last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones",
                      "Garcia", "Miller", "Davis", "Rodriguez", "Martinez",
                      "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson",
                      "Thomas", "Taylor", "Moore", "Jackson", "Martin", "Lee",
                      "Perez", "Thompson", "White", "Harris"]

        employee_ids = random.sample(range(1000, 2500), len(first_names))

        countries = {"USA": 1, "India": 2, "China": 3, "France": 4, "Germany": 5}
        country_names = list(countries.keys())
        country_codes = list(countries.values())

        departments = ["Production", "R&D", "Marketing", "HR",
                       "Finance", "Engineering", "Managerial"]

        for f_name in first_names:
            l_name = last_names.pop()
            # email = (l_name + f_name[0]).lower() + "@job.com"
            country_id = random.choice(country_codes)
            dept = random.choice(departments)
            employee_id = employee_ids.pop()
            query.addBindValue(employee_id)
            query.addBindValue(f_name)
            query.addBindValue(l_name)
            # query.addBindValue(email)
            query.addBindValue(dept)
            query.addBindValue(country_id)
            query.exec_()

        # Create the second table, countries
        country_query = QSqlQuery()
        country_query.exec_("""CREATE TABLE countries (
                             id INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL,
                             country VARCHAR(20) NOT NULL)""")

        country_query.prepare("INSERT INTO countries (country) VALUES (?)")

        for name in country_names:
            country_query.addBindValue(name)
            country_query.exec_()

        print("[INFO] Database successfully created.")

        # sys.exit(0)

    def createConnection(self):

        if QSqlDatabase.contains("qt_sql_default_connection"):
            return True  # Ya hay una conexión activa, no es necesario volver a conectar

        database2 = QSqlDatabase.addDatabase("QSQLITE")  # SQLite version 3
        database2.setDatabaseName("files/accounts.db")
        # database.setDatabaseName(r"\\gfa60005\intercambios\AlRa\db\accounts.db")

        if not database2.open():
            print("Unable to open data source file.")  # esto ocurre cuando el directorio en sí no existe
            sys.exit(1)  # Código de error 1 - indica error

        # Verificar si las tablas necesarias existen en la base de datos
        tables_needed = {'accounts', 'countries'}
        tables_not_found = tables_needed - set(database2.tables())
        if tables_not_found:
            QMessageBox.critical(None, 'Error',
                                 f'The following tables tables are missing from the database: {tables_not_found}')
            sys.exit(1)  # Código de error 1 - indica error

        return True  # La conexión se ha establecido correctamente


        # database2 = QSqlDatabase.addDatabase("QSQLITE")  # SQLite version 3
        # database2.setDatabaseName("files/accounts.db")
        # #database.setDatabaseName(r"\\gfa60005\intercambios\AlRa\db\accounts.db")
        #
        # if not database2.open():
        #     print("Unable to open data source file.")  # this happens when the directory itself does not exist
        #     sys.exit(1)  # Error code 1 - signifies error
        #
        # # Check if the tables we need exist in the database
        # tables_needed = {'accounts', 'countries'}
        # tables_not_found = tables_needed - set(database2.tables())
        # if tables_not_found:
        #     QMessageBox.critical(None, 'Error',
        #                          f'The following tables tables are missing from the database: {tables_not_found}')
        #     sys.exit(1)  # Error code 1 - signifies error

    def createTable(self):
        """
        Set up the model, headers and populate the model.
        """
        self.model = QSqlRelationalTableModel()
        self.model.setTable('accounts')
        self.model.setRelation(self.model.fieldIndex('country_id'), QSqlRelation('countries', 'id', 'country'))

        self.model.setHeaderData(self.model.fieldIndex('id'), Qt.Horizontal, "ID")
        self.model.setHeaderData(self.model.fieldIndex('employee_id'), Qt.Horizontal, "Employee ID")
        self.model.setHeaderData(self.model.fieldIndex('first_name'), Qt.Horizontal, "First")
        self.model.setHeaderData(self.model.fieldIndex('last_name'), Qt.Horizontal, "Last")
        # self.model.setHeaderData(self.model.fieldIndex('email'), Qt.Horizontal, "E-mail")
        self.model.setHeaderData(self.model.fieldIndex('department'), Qt.Horizontal, "Dept.")
        self.model.setHeaderData(self.model.fieldIndex('country_id'), Qt.Horizontal, "Country")

        # Populate the model with data
        self.model.select()

    def setupWidgets(self):
        """
        Create instances of widgets, the table view and set layouts.
        """
        icons_path = "icons"

        title = QLabel("Add your new comment here ==>")
        title.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        title.setStyleSheet("font: bold 20px")

        add_product_button = QPushButton("Add Employee")
        add_product_button.setIcon(QIcon(os.path.join(icons_path, "add_user.png")))
        add_product_button.setStyleSheet("padding: 10px")
        add_product_button.clicked.connect(self.addItem)

        del_product_button = QPushButton("Delete")
        del_product_button.setIcon(QIcon(os.path.join(icons_path, "trash_can.png")))
        del_product_button.setStyleSheet("padding: 10px")
        del_product_button.clicked.connect(self.deleteItem)

        # Set up sorting combobox
        sorting_options = ["Sort by ID", "Sort by Employee ID", "Sort by First Name",
                           "Sort by Last Name", "Sort by Department", "Sort by Country"]
        sort_name_cb = QComboBox()
        sort_name_cb.addItems(sorting_options)
        sort_name_cb.currentTextChanged.connect(self.setSortingOrder)

        buttons_h_box = QHBoxLayout()
        buttons_h_box.addWidget(add_product_button)
        buttons_h_box.addWidget(del_product_button)
        buttons_h_box.addStretch()
        buttons_h_box.addWidget(sort_name_cb)

        # Widget to contain editing buttons
        edit_buttons = QWidget()
        edit_buttons.setLayout(buttons_h_box)

        # Create table view and set model
        self.table_view = QTableView()
        self.table_view.setModel(self.model)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_view.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_view.setSelectionMode(QTableView.SingleSelection)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)

        # Instantiate the delegate
        delegate = QSqlRelationalDelegate(self.table_view)
        self.table_view.setItemDelegate(delegate)

        # Main layout
        main_h = QHBoxLayout()

        main_v_box = QVBoxLayout()
        main_v_box.addWidget(title, Qt.AlignLeft)
        main_v_box.addWidget(edit_buttons)
        main_v_box.addWidget(self.table_view)

        main_v_box_der = QVBoxLayout()
        self.comments = QTextEdit()
        main_v_box_der.addWidget(self.comments)

        main_h.addLayout(main_v_box)
        main_h.addLayout(main_v_box_der)

        self.setLayout(main_h)

    def addItem(self):
        """
        Add a new record to the last row of the table.
        """
        last_row = self.model.rowCount()
        self.model.insertRow(last_row)

        id = 0
        query = QSqlQuery()
        query.exec_("SELECT MAX (id) FROM accounts")
        if query.next():
            print(query.value(0))
            id = int(query.value(0))

    def deleteItem(self):
        """
        Delete an entire row from the table.
        """
        current_item = self.table_view.selectedIndexes()
        for index in current_item:
            self.model.removeRow(index.row())
        self.model.select()

    def setSortingOrder(self, text):
        """
        Sort the rows in table.
        """
        if text == "Sort by ID":
            self.model.setSort(self.model.fieldIndex('id'), Qt.AscendingOrder)
        elif text == "Sort by Employee ID":
            self.model.setSort(self.model.fieldIndex('employee_id'), Qt.AscendingOrder)
        elif text == "Sort by First Name":
            self.model.setSort(self.model.fieldIndex('first_name'), Qt.AscendingOrder)
        elif text == "Sort by Last Name":
            self.model.setSort(self.model.fieldIndex('last_name'), Qt.AscendingOrder)
        elif text == "Sort by Department":
            self.model.setSort(self.model.fieldIndex('department'), Qt.AscendingOrder)
        elif text == "Sort by Country":
            self.model.setSort(self.model.fieldIndex('country'), Qt.AscendingOrder)

        self.model.select()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setStyleSheet(style_sheet)
    #form = MainGUI()
    form = MainWindow()
    sys.exit(app.exec_())



