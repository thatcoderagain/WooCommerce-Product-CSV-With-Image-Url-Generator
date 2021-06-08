import os, sys
import json
import pandas
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from google.oauth2 import service_account
from googleapiclient.discovery import build

BASE_IMAGE_PATH_URL = 'https://homefabindia.com/wp-content/uploads/images/products/curtains/'
SKU_ID_EXTRAS = ['5F','6F','7F','8F','9F','Setof2']
IMAGES_FOLDER = 'Converted images'
SKU_FILE = 'SKU_List.txt'
EXCEL_FILE = 'Products.xlsx'
SKU_READ_METHOD = 'readFromLocal'
EXPORT_METHOD = 'exportToLocal'

SERVICE_ACCOUNT_FILE = './docs-316004-54c2dd979ce3.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDENTIALS = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
SAMPLE_SPREADSHEET_ID = '1S0vxL_-7bGZ64I4b3s86L4T0G8PNf6qoRxr4YxwUsd4'
READ_RANGE_NAME = 'Data!C2:C1000'
WRITE_RANGE_NAME_1 = 'Compiled!AD3'
WRITE_RANGE_NAME_2 = 'Compiled!AN3'

def createConfig():
    configDict = {}
    configDict['BASE_IMAGE_PATH_URL'] = BASE_IMAGE_PATH_URL
    configDict['SKU_ID_EXTRAS'] = sorted(SKU_ID_EXTRAS)
    configDict['IMAGES_FOLDER'] = IMAGES_FOLDER
    configDict['SKU_FILE'] = SKU_FILE
    configDict['EXCEL_FILE'] = EXCEL_FILE
    configDict['SKU_READ_METHOD'] = SKU_READ_METHOD
    configDict['EXPORT_METHOD'] = EXPORT_METHOD
    configDict['SAMPLE_SPREADSHEET_ID'] = SAMPLE_SPREADSHEET_ID
    configDict['READ_RANGE_NAME'] = READ_RANGE_NAME
    configDict['WRITE_RANGE_NAME_1'] = WRITE_RANGE_NAME_1
    configDict['WRITE_RANGE_NAME_2'] = WRITE_RANGE_NAME_2
    return configDict

def config():
    if (os.path.exists('app.config') == False):
        configDict = createConfig()
        json_object = json.dumps(configDict, indent = 4)
        file = open('app.config', 'w')
        file.write(json_object)
    else:
        import ast
        configDict = {}
        with open("app.config", "r") as data:
            configDict = {**configDict, **ast.literal_eval(data.read())}
    return dict(configDict)


def updateConfig(configDict):
    os.rename('app.config','app.config.bak')
    json_object = json.dumps(configDict, indent = 4)
    file = open('app.config', 'w')
    file.write(json_object)
    if (os.path.exists('app.config') == True):
        os.remove('app.config.bak')


def readExcel(path, sheetName='Sheet1'):
    path = path.replace('\\', '/')
    excel = pandas.read_excel(path, sheet_name=sheetName)
    data = list()
    data = data + [list(map(lambda x: x, excel.columns))]
    for index, row in excel.iterrows():
        data = data + [list(map(lambda x: x, row))]
    return data


def writeToSpreadSheet(range, data):
    global SERVICE_ACCOUNT_FILE, SCOPES, CREDENTIALS, SAMPLE_SPREADSHEET_ID, WRITE_RANGE_NAME_1, WRITE_RANGE_NAME_2
    service = build('sheets', 'v4', credentials=CREDENTIALS)
    value_range_body = {}
    value_range_body['values'] = data
    response = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range,
                                                    valueInputOption="USER_ENTERED", body=value_range_body).execute()


def clearSheet(sheetName):
    global CREDENTIALS
    service = build('sheets', 'v4', credentials=CREDENTIALS)
    rangeAll = '{0}!A1:ZZ'.format(sheetName)
    response = service.spreadsheets().values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, body={},
                                                    range='{0}!A1:Z'.format(sheetName)).execute()


def readSKUIdsFromSpreadSheet():
    global SERVICE_ACCOUNT_FILE, SCOPES, CREDENTIALS, SAMPLE_SPREADSHEET_ID, READ_RANGE_NAME
    service = build('sheets', 'v4', credentials=CREDENTIALS)
    response = service.spreadsheets().values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=READ_RANGE_NAME).execute()
    values = response.get('values', [])
    return list(map(lambda x: (x[0]), values))


def writeImagesURLToSpreadSheet(dateColumn1, dateColumn2):
    global SERVICE_ACCOUNT_FILE, SCOPES, CREDENTIALS, SAMPLE_SPREADSHEET_ID, WRITE_RANGE_NAME_1, WRITE_RANGE_NAME_2
    service = build('sheets', 'v4', credentials=CREDENTIALS)
    value_range_body = {}
    value_range_body['values'] = dateColumn1
    response = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=WRITE_RANGE_NAME_1,
                                                    valueInputOption="USER_ENTERED", body=value_range_body).execute()
    value_range_body = {}
    value_range_body['values'] = dateColumn2
    response = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=WRITE_RANGE_NAME_2,
                                                    valueInputOption="USER_ENTERED", body=value_range_body).execute()


def customFilter(string, targets):
    for target in targets:
        string = string[len(target):] if string.startswith(target) else string
        string = string[:len(string)-len(target)] if string.endswith(target) else string
    return string


def generateURLList():
    if (os.path.isdir(IMAGES_FOLDER) == False):
        return [False, "Oops! Images folder not found"]

    listOfFiles = os.listdir(IMAGES_FOLDER)
    images = filter(lambda x: x.endswith('.jpg'), listOfFiles)
    images = sorted(images)
    images = list(map(lambda x: BASE_IMAGE_PATH_URL+x, images))

    if (len(images) == 0):
        return [False, "Oops! Images not found in the selected folder"]

    if SKU_READ_METHOD == 'readFromLocal':
        if (os.path.exists(SKU_FILE) == False):
            return [False, "Oops! SKU file not found"]
        skuFile = open(SKU_FILE, 'r')
        skuIds = skuFile.readlines()
    elif SKU_READ_METHOD == 'readFromGoogleSpreadSheet':
        skuIds = readSKUIdsFromSpreadSheet()
    elif SKU_READ_METHOD == 'readExcelAndExportProductToGoogleSpreadSheet':
        if (os.path.exists(EXCEL_FILE) == False):
            return [False, "Oops! Product excel file not found"]
        data = readExcel(EXCEL_FILE, 'Data')
        response = clearSheet('Data')
        response = writeToSpreadSheet('Data!A1', data)
        skuIds = readSKUIdsFromSpreadSheet()

    skuIds = map(lambda x: (x.strip()), skuIds)
    skuIds = list(map(lambda x: (customFilter(x, SKU_ID_EXTRAS)), skuIds))

    if (len(skuIds) == 0):
        return [False, "Oops! No SKU ids found for the products"]

    imagesSet = {}
    for skuId in skuIds:
        imagesSet[skuId] = list()

    for skuId in skuIds:
        for image in images:
            ls = []
            if (skuId in image):
                ls.append(image)
            if (len(ls) > 0):
                imagesSet[skuId] = imagesSet[skuId] + ls
                imagesSet[skuId] = list(set(imagesSet[skuId]))

    if SKU_READ_METHOD == 'readFromLocal' or EXPORT_METHOD == 'exportToLocal':
        AD3 = open('AD3.csv', 'w')
        AN3 = open('AN3.csv', 'w')
        for skuId in skuIds:
            AD3.write(",".join(imagesSet[skuId][:1])+'\n')
            AN3.write(",".join(imagesSet[skuId][1:])+'\n')
    elif SKU_READ_METHOD == 'readExcelAndExportProductToGoogleSpreadSheet' or EXPORT_METHOD == 'exportUrlToGoogleSpreadSheet':
        AD3 = []
        AN3 = []
        for skuId in skuIds:
            AD3.append([",".join(imagesSet[skuId][:1])])
            AN3.append([",".join(imagesSet[skuId][1:])])
        AD3 = AD3 + [['']]*(1000 - len(AD3))
        AN3 = AN3 + [['']]*(1000 - len(AN3))
        writeImagesURLToSpreadSheet(AD3, AN3)

    updateConfig(createConfig())
    return [True, "Boom! All done Khushi Goyal :*"]


class Widgets(QWidget):
    def __init__(self, **kwargs):
        super(Widgets, self).__init__()

        self.setWindowTitle("Khushi Tool")
        self.setGeometry(100,100,900,250)
        self.move(200,200)

        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        self.horizontalLayoutBaseImageUrl = QHBoxLayout()
        self.labelBaseImageUrl = QLabel()
        self.labelBaseImageUrl.setText("Base image url ")
        self.horizontalLayoutBaseImageUrl.addWidget(self.labelBaseImageUrl)
        self.lineEditBaseImageUrl = QLineEdit()
        self.lineEditBaseImageUrl.setFixedWidth(712)
        self.lineEditBaseImageUrl.setText(BASE_IMAGE_PATH_URL)
        self.horizontalLayoutBaseImageUrl.addWidget(self.lineEditBaseImageUrl)
        self.verticalLayout.addLayout(self.horizontalLayoutBaseImageUrl)

        self.horizontalLayoutSKUVariations = QHBoxLayout()
        self.labelSkuIdVariations = QLabel()
        self.labelSkuIdVariations.setText("SKU variations ")
        self.horizontalLayoutSKUVariations.addWidget(self.labelSkuIdVariations)
        self.lineEditSkuIdVariations = QLineEdit()
        self.lineEditSkuIdVariations.setFixedWidth(712)
        self.lineEditSkuIdVariations.setText(",".join(SKU_ID_EXTRAS))
        self.horizontalLayoutSKUVariations.addWidget(self.lineEditSkuIdVariations)
        self.verticalLayout.addLayout(self.horizontalLayoutSKUVariations)

        self.horizontalLayoutImagesFolderName = QHBoxLayout()
        self.labelImagesFolder = QLabel()
        self.labelImagesFolder.setText("Images Folder ")
        self.horizontalLayoutImagesFolderName.addWidget(self.labelImagesFolder)
        self.lineEditImageFolder = QLineEdit()
        self.lineEditImageFolder.setFixedWidth(500)
        self.lineEditImageFolder.setText(IMAGES_FOLDER)
        self.horizontalLayoutImagesFolderName.addWidget(self.lineEditImageFolder)
        self.buttonSkuFilePicker = QPushButton("Browse (Images) ")
        self.buttonSkuFilePicker.setFixedWidth(205)
        self.buttonSkuFilePicker.clicked.connect(self.onButtonImageFolderPickerClick)
        self.horizontalLayoutImagesFolderName.addWidget(self.buttonSkuFilePicker)
        self.verticalLayout.addLayout(self.horizontalLayoutImagesFolderName)

        self.horizontalLayoutSkuReadMethod = QHBoxLayout()
        self.buttonGroupSkuRead = QButtonGroup()
        self.labelSkuReadMethod = QLabel()
        self.labelSkuReadMethod.setText("SKU Read Method")
        self.radiobuttonReadSKULocal = QRadioButton("Local File")
        self.radiobuttonReadSKULocal.setFixedWidth(170)
        self.radiobuttonReadSKULocal.method = "readFromLocal"
        self.radiobuttonReadSKULocal.toggled.connect(self.onSKUReadMethodToggled)
        self.radiobuttonReadSKUGoogleSpreadSheet = QRadioButton("Google SpreadSheet")
        self.radiobuttonReadSKUGoogleSpreadSheet.setFixedWidth(200)
        self.radiobuttonReadSKUGoogleSpreadSheet.method = "readFromGoogleSpreadSheet"
        self.radiobuttonReadSKUGoogleSpreadSheet.toggled.connect(self.onSKUReadMethodToggled)
        self.radiobuttonUploadProductToGoogleSpreadSheet = QRadioButton("Import Product + Upload to SpreadSheet")
        self.radiobuttonUploadProductToGoogleSpreadSheet.setFixedWidth(330)
        self.radiobuttonUploadProductToGoogleSpreadSheet.method = "readExcelAndExportProductToGoogleSpreadSheet"
        self.radiobuttonUploadProductToGoogleSpreadSheet.toggled.connect(self.onSKUReadMethodToggled)
        self.buttonGroupSkuRead.addButton(self.radiobuttonReadSKULocal)
        self.buttonGroupSkuRead.addButton(self.radiobuttonReadSKUGoogleSpreadSheet)
        self.buttonGroupSkuRead.addButton(self.radiobuttonUploadProductToGoogleSpreadSheet)
        self.horizontalLayoutSkuReadMethod.addWidget(self.labelSkuReadMethod)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonReadSKULocal)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonReadSKUGoogleSpreadSheet)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonUploadProductToGoogleSpreadSheet)
        self.verticalLayout.addLayout(self.horizontalLayoutSkuReadMethod)


        self.horizontalLayoutSkuReadMethod = QHBoxLayout()
        self.buttonGroupExportUrl = QButtonGroup()
        self.labelSkuReadMethod = QLabel()
        self.labelSkuReadMethod.setText("Export Method")
        self.radiobuttonExportUrlLocal = QRadioButton("Generate .csv files")
        self.radiobuttonExportUrlLocal.method = "exportToLocal"
        self.radiobuttonExportUrlLocal.setFixedWidth(170)
        self.radiobuttonExportUrlLocal.toggled.connect(self.onExportMethodToggled)
        self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet = QRadioButton("Export Urls to SpreadSheet")
        self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet.setFixedWidth(535)
        self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet.method = "exportUrlToGoogleSpreadSheet"
        self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet.toggled.connect(self.onExportMethodToggled)
        self.buttonGroupExportUrl.addButton(self.radiobuttonExportUrlLocal)
        self.buttonGroupExportUrl.addButton(self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet)
        self.horizontalLayoutSkuReadMethod.addWidget(self.labelSkuReadMethod)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonExportUrlLocal)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet)
        self.labelSkuReadMethod.hide()
        self.radiobuttonExportUrlLocal.hide()
        self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet.hide()
        self.verticalLayout.addLayout(self.horizontalLayoutSkuReadMethod)

        self.horizontalLayoutSKUFilename = QHBoxLayout()
        self.labelSkuIdFilename = QLabel()
        self.labelSkuIdFilename.setText("SKU Filename ")
        self.horizontalLayoutSKUFilename.addWidget(self.labelSkuIdFilename)
        self.lineEditSkuIdFilename = QLineEdit()
        self.lineEditSkuIdFilename.setFixedWidth(500)
        self.lineEditSkuIdFilename.setText(SKU_FILE)
        self.horizontalLayoutSKUFilename.addWidget(self.lineEditSkuIdFilename)
        self.buttonSkuFilePicker = QPushButton("Browse (SKU List.txt) ")
        self.buttonSkuFilePicker.setFixedWidth(205)
        self.buttonSkuFilePicker.clicked.connect(self.onButtonSkuFilePickerClick)
        self.horizontalLayoutSKUFilename.addWidget(self.buttonSkuFilePicker)
        self.verticalLayout.addLayout(self.horizontalLayoutSKUFilename)

        self.horizontalLayoutExcelFilePicker = QHBoxLayout()
        self.labelExcelFilename = QLabel()
        self.labelExcelFilename.setText("Product File ")
        self.horizontalLayoutExcelFilePicker.addWidget(self.labelExcelFilename)
        self.lineEditExcelFilename = QLineEdit()
        self.lineEditExcelFilename.setFixedWidth(500)
        self.lineEditExcelFilename.setText(EXCEL_FILE)
        self.horizontalLayoutExcelFilePicker.addWidget(self.lineEditExcelFilename)
        self.buttonExcelFilename = QPushButton("Browse (Products.xlsx)")
        self.buttonExcelFilename.setFixedWidth(205)
        self.buttonExcelFilename.clicked.connect(self.onButtonExcelFilePickerClick)
        self.horizontalLayoutExcelFilePicker.addWidget(self.buttonExcelFilename)
        self.verticalLayout.addLayout(self.horizontalLayoutExcelFilePicker)

        self.buttonSubmit = QPushButton("Generate")
        self.buttonSubmit.clicked.connect(self.onButtonSubmitClick)
        self.verticalLayout.addWidget(self.buttonSubmit)
        self.setLayout(self.verticalLayout)

        if SKU_READ_METHOD == 'readFromLocal':
            self.radiobuttonReadSKULocal.setChecked(True)
        elif SKU_READ_METHOD == 'readFromGoogleSpreadSheet':
            self.radiobuttonReadSKUGoogleSpreadSheet.setChecked(True)
            if EXPORT_METHOD == 'exportToLocal':
                self.radiobuttonExportUrlLocal.setChecked(True)
            elif EXPORT_METHOD == 'exportUrlToGoogleSpreadSheet':
                self.radiobuttonUploadProductToGoogleSpreadSheet.setChecked(True)
        elif SKU_READ_METHOD == 'readExcelAndExportProductToGoogleSpreadSheet':
            self.radiobuttonUploadProductToGoogleSpreadSheet.setChecked(True)

    def onSKUReadMethodToggled(self):
        global SKU_READ_METHOD, EXPORT_METHOD
        radiobuttonReadSKU = self.sender()
        if radiobuttonReadSKU.method in ['readFromLocal', 'readFromGoogleSpreadSheet', 'readExcelAndExportProductToGoogleSpreadSheet']:
            SKU_READ_METHOD = radiobuttonReadSKU.method
            if radiobuttonReadSKU.isChecked():
                if radiobuttonReadSKU.method == 'readFromLocal':
                    self.labelSkuIdFilename.setHidden(False)
                    self.lineEditSkuIdFilename.show()
                    self.buttonSkuFilePicker.show()
                    self.labelSkuReadMethod.hide()
                    self.radiobuttonExportUrlLocal.hide()
                    self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet.hide()
                    self.labelExcelFilename.setHidden(True)
                    self.lineEditExcelFilename.hide()
                    self.buttonExcelFilename.hide()
                elif radiobuttonReadSKU.method == 'readFromGoogleSpreadSheet':
                    self.labelSkuIdFilename.setHidden(True)
                    self.lineEditSkuIdFilename.hide()
                    self.buttonSkuFilePicker.hide()
                    self.labelSkuReadMethod.show()
                    self.radiobuttonExportUrlLocal.show()
                    self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet.show()
                    self.labelExcelFilename.setHidden(True)
                    self.lineEditExcelFilename.hide()
                    self.buttonExcelFilename.hide()
                elif radiobuttonReadSKU.method == 'readExcelAndExportProductToGoogleSpreadSheet':
                    self.labelSkuIdFilename.setHidden(True)
                    self.lineEditSkuIdFilename.hide()
                    self.buttonSkuFilePicker.hide()
                    self.labelSkuReadMethod.hide()
                    self.radiobuttonExportUrlLocal.hide()
                    self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet.hide()
                    self.labelExcelFilename.setHidden(False)
                    self.lineEditExcelFilename.show()
                    self.buttonExcelFilename.show()

    def onExportMethodToggled(self):
        global SKU_READ_METHOD, EXPORT_METHOD
        radiobuttonExportUrl = self.sender()
        if radiobuttonExportUrl.method in ['exportToLocal', 'exportUrlToGoogleSpreadSheet']:
            EXPORT_METHOD = radiobuttonExportUrl.method

    def onButtonSkuFilePickerClick(self):
        global SKU_FILE
        filepath = self.filePicker()
        if (filepath):
            SKU_FILE = filepath
            self.lineEditSkuIdFilename.setText(SKU_FILE)

    def onButtonExcelFilePickerClick(self):
        global EXCEL_FILE
        filepath = self.filePicker()
        if (filepath):
            if (filepath.endswith('xlsx') == True):
                EXCEL_FILE = filepath
                self.lineEditExcelFilename.setText(EXCEL_FILE)
            else:
                self.showDialog('Invalid file. Please select a valid excel file.')

    def onButtonImageFolderPickerClick(self):
        filepath = self.folderPicker()
        if (filepath):
            IMAGES_FOLDER = filepath
            self.lineEditImageFolder.setText(IMAGES_FOLDER)

    def onButtonSubmitClick(self):
        global BASE_IMAGE_PATH_URL, SKU_ID_EXTRAS, IMAGES_FOLDER, SKU_FILE, SKU_READ_METHOD, EXPORT_METHOD, EXCEL_FILE
        msg = ""
        self.buttonSubmit.setEnabled(False)
        if (self.lineEditBaseImageUrl.text().startswith('https://')):
            BASE_IMAGE_PATH_URL = self.lineEditBaseImageUrl.text().strip()
            if (BASE_IMAGE_PATH_URL.endswith('/') == False):
                BASE_IMAGE_PATH_URL = BASE_IMAGE_PATH_URL+'/'
        else:
            msg = "Invalid base URL provided"

        if (len(self.lineEditImageFolder.text()) > 0):
            IMAGES_FOLDER = self.lineEditImageFolder.text()
        else:
            msg = "No images folder provided"

        SKU_ID_EXTRAS = list(map(lambda x: (x.strip()), list(set(self.lineEditSkuIdVariations.text().strip().split(',')))))

        if (SKU_READ_METHOD == 'readFromLocal'):
            if (len(self.lineEditSkuIdFilename.text()) > 0):
                SKU_FILE = self.lineEditSkuIdFilename.text()
            else:
                msg = "No SKU filename provided"
        elif (SKU_READ_METHOD == 'readFromGoogleSpreadSheet'):
            pass
        elif (SKU_READ_METHOD == 'readExcelAndExportProductToGoogleSpreadSheet'):
            if (len(self.lineEditExcelFilename.text()) > 0):
                SKU_FILE = self.lineEditExcelFilename.text()
            else:
                msg = "No SKU filename provided"

        if (len(msg) > 0):
            self.showDialog(msg)
            return

        result = generateURLList()
        self.showDialog(result[1], result[0])
        self.buttonSubmit.setEnabled(True)

    def showDialog(self, msgText = "Something went wrong!", status = False):
        msg = QMessageBox()
        msg.setWindowTitle("Success" if status else "Oops!!!")
        msg.setIcon(QMessageBox.Information if status else QMessageBox.Critical)
        msg.setText("Process completed" if status else "Process failed")
        msg.setDetailedText(f"The details are as follows:\n{msgText}")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.exec_()
        self.buttonSubmit.setEnabled(True)

    def folderPicker(self):
        directory = QFileDialog.getExistingDirectory(self, "Select a folder", "./")
        if (len(directory) > 0):
            return directory
        else:
            return False

    def filePicker(self):
        filepath = QFileDialog.getOpenFileName(self, "Select a file", "./", "All Files (*);;") #
        if (len(filepath[0]) > 0):
            return filepath[0]
        else:
            return False


def window():
    app = QApplication([])

    widget = Widgets()
    widget.show()

    app.exec_()


if __name__ == '__main__':
    config = config()
    BASE_IMAGE_PATH_URL = config['BASE_IMAGE_PATH_URL']
    SKU_ID_EXTRAS = config['SKU_ID_EXTRAS']
    IMAGES_FOLDER = config['IMAGES_FOLDER']
    SKU_FILE = config['SKU_FILE']
    EXCEL_FILE = config['EXCEL_FILE']
    SKU_READ_METHOD = config['SKU_READ_METHOD']
    EXPORT_METHOD = config['EXPORT_METHOD']
    SAMPLE_SPREADSHEET_ID = config['SAMPLE_SPREADSHEET_ID']
    READ_RANGE_NAME = config['READ_RANGE_NAME']
    WRITE_RANGE_NAME_1 = config['WRITE_RANGE_NAME_1']
    WRITE_RANGE_NAME_2 = config['WRITE_RANGE_NAME_2']
    window()
