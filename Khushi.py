import os, sys
import json
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from google.oauth2 import service_account
from googleapiclient.discovery import build

BASE_IMAGE_PATH_URL = 'https://homefabindia.com/wp-content/uploads/images/products/curtains/'
SKU_ID_EXTRAS = ['5F','6F','7F','8F','9F','Setof2']
IMAGES_FOLDER = 'Converted images'
SKU_FILE = 'SKU_List.txt'
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


def customFilter(x, targets):
    for target in targets:
        x = x.replace(target, '')
    return x


def generateURLList():
    global configDict
    if (os.path.isdir(IMAGES_FOLDER) == False):
        return [False, "Oops! Images folder not found"]

    listOfFiles = os.listdir(IMAGES_FOLDER)
    images = filter(lambda x: x.endswith('.jpg'), listOfFiles)
    images = sorted(images)
    images = list(map(lambda x: BASE_IMAGE_PATH_URL+x, images))

    if (len(images) == 0):
        return [False, "Oops! Images not found in the selected folder"]

    if SKU_READ_METHOD == 'readFromLocal':
        skuFile = open(SKU_FILE, 'r')
        skuIds = skuFile.readlines()
    elif SKU_READ_METHOD == 'readFromGoogleSpreadSheet':
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

    if EXPORT_METHOD == 'exportToLocal':
        AD3 = open('AD3.csv', 'w')
        AN3 = open('AN3.csv', 'w')
        for skuId in skuIds:
            AD3.write(",".join(imagesSet[skuId][:1])+'\n')
            AN3.write(",".join(imagesSet[skuId][1:])+'\n')
    elif EXPORT_METHOD == 'exportToGoogleSpreadSheet':
        AD3 = []
        AN3 = []
        for skuId in skuIds:
            AD3.append([",".join(imagesSet[skuId][:1])])
            AN3.append([",".join(imagesSet[skuId][1:])])
        writeImagesURLToSpreadSheet(AD3, AN3)

    updateConfig(createConfig())
    return [True, "Boom! All done Khushi Goyal :*"]


class Widgets(QWidget):
    def __init__(self, **kwargs):
        super(Widgets, self).__init__()

        self.setWindowTitle("Khushi Tool")
        self.setGeometry(100,100,800,250)
        self.move(200,200)

        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        self.horizontalLayoutBaseImageUrl = QHBoxLayout()
        self.labelBaseImageUrl = QLabel()
        self.labelBaseImageUrl.setText("Base image url ")
        self.horizontalLayoutBaseImageUrl.addWidget(self.labelBaseImageUrl)
        self.lineEditBaseImageUrl = QLineEdit()
        self.lineEditBaseImageUrl.setFixedWidth(562)
        self.lineEditBaseImageUrl.setText(BASE_IMAGE_PATH_URL)
        self.horizontalLayoutBaseImageUrl.addWidget(self.lineEditBaseImageUrl)
        self.verticalLayout.addLayout(self.horizontalLayoutBaseImageUrl)

        self.horizontalLayoutSKUVariations = QHBoxLayout()
        self.labelSkuIdVariations = QLabel()
        self.labelSkuIdVariations.setText("SKU variations ")
        self.horizontalLayoutSKUVariations.addWidget(self.labelSkuIdVariations)
        self.lineEditSkuIdVariations = QLineEdit()
        self.lineEditSkuIdVariations.setFixedWidth(562)
        self.lineEditSkuIdVariations.setText(",".join(SKU_ID_EXTRAS))
        self.horizontalLayoutSKUVariations.addWidget(self.lineEditSkuIdVariations)
        self.verticalLayout.addLayout(self.horizontalLayoutSKUVariations)

        self.horizontalLayoutImagesFolderName = QHBoxLayout()
        self.labelImagesFolder = QLabel()
        self.labelImagesFolder.setText("Images Folder ")
        self.horizontalLayoutImagesFolderName.addWidget(self.labelImagesFolder)
        self.lineEditImageFolder = QLineEdit()
        self.lineEditImageFolder.setFixedWidth(350)
        self.lineEditImageFolder.setText(IMAGES_FOLDER)
        self.horizontalLayoutImagesFolderName.addWidget(self.lineEditImageFolder)
        self.buttonSKUFilePicker = QPushButton("Browse (Images) ")
        self.buttonSKUFilePicker.clicked.connect(self.onButtonImageFolderPickerClick)
        self.horizontalLayoutImagesFolderName.addWidget(self.buttonSKUFilePicker)
        self.verticalLayout.addLayout(self.horizontalLayoutImagesFolderName)

        self.horizontalLayoutSkuReadMethod = QHBoxLayout()
        self.buttonGroupSkuRead = QButtonGroup()
        self.labelSkuReadMethod = QLabel()
        self.labelSkuReadMethod.setText("SKU Read Method")
        self.radiobuttonReadSKULocal = QRadioButton("Local File")
        self.radiobuttonReadSKULocal.setFixedWidth(200)
        self.radiobuttonReadSKULocal.method = "readFromLocal"
        self.radiobuttonReadSKULocal.toggled.connect(self.onSKUReadMethodToggled)
        self.radiobuttonReadSKUGoogleSpreadSheet = QRadioButton("Google SpreadSheet")
        self.radiobuttonReadSKUGoogleSpreadSheet.setFixedWidth(350)
        self.radiobuttonReadSKUGoogleSpreadSheet.method = "readFromGoogleSpreadSheet"
        self.radiobuttonReadSKUGoogleSpreadSheet.toggled.connect(self.onSKUReadMethodToggled)
        self.radiobuttonReadSKULocal.setChecked(True)
        self.buttonGroupSkuRead.addButton(self.radiobuttonReadSKULocal)
        self.buttonGroupSkuRead.addButton(self.radiobuttonReadSKUGoogleSpreadSheet)
        self.horizontalLayoutSkuReadMethod.addWidget(self.labelSkuReadMethod)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonReadSKULocal)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonReadSKUGoogleSpreadSheet)
        self.verticalLayout.addLayout(self.horizontalLayoutSkuReadMethod)


        self.horizontalLayoutSkuReadMethod = QHBoxLayout()
        self.buttonGroupExportUrl = QButtonGroup()
        self.labelSkuReadMethod = QLabel()
        self.labelSkuReadMethod.setText("Export Method")
        self.radiobuttonExportUrlLocal = QRadioButton("Generate .csv files")
        self.radiobuttonExportUrlLocal.method = "exportToLocal"
        self.radiobuttonExportUrlLocal.setFixedWidth(200)
        self.radiobuttonExportUrlLocal.toggled.connect(self.onExportMethodToggled)
        self.radiobuttonExportUrlGoogleSpreadSheet = QRadioButton("Upload to Google SpreadSheet")
        self.radiobuttonExportUrlGoogleSpreadSheet.setFixedWidth(350)
        self.radiobuttonExportUrlGoogleSpreadSheet.method = "exportToGoogleSpreadSheet"
        self.radiobuttonExportUrlGoogleSpreadSheet.toggled.connect(self.onExportMethodToggled)
        self.buttonGroupExportUrl.addButton(self.radiobuttonExportUrlLocal)
        self.buttonGroupExportUrl.addButton(self.radiobuttonExportUrlGoogleSpreadSheet)
        self.horizontalLayoutSkuReadMethod.addWidget(self.labelSkuReadMethod)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonExportUrlLocal)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonExportUrlGoogleSpreadSheet)
        self.radiobuttonExportUrlLocal.setChecked(True)
        self.labelSkuReadMethod.hide()
        self.radiobuttonExportUrlLocal.hide()
        self.radiobuttonExportUrlGoogleSpreadSheet.hide()
        self.verticalLayout.addLayout(self.horizontalLayoutSkuReadMethod)

        self.horizontalLayoutSKUFilename = QHBoxLayout()
        self.labelSkuIdFilename = QLabel()
        self.labelSkuIdFilename.setText("SKU Filename ")
        self.horizontalLayoutSKUFilename.addWidget(self.labelSkuIdFilename)
        self.lineEditSkuIdFilename = QLineEdit()
        self.lineEditSkuIdFilename.setFixedWidth(350)
        self.lineEditSkuIdFilename.setText(SKU_FILE)
        self.horizontalLayoutSKUFilename.addWidget(self.lineEditSkuIdFilename)
        self.buttonSKUFilePicker = QPushButton("Browse (SKU List.txt) ")
        self.buttonSKUFilePicker.clicked.connect(self.onButtonSKUFilePickerClick)
        self.horizontalLayoutSKUFilename.addWidget(self.buttonSKUFilePicker)
        self.verticalLayout.addLayout(self.horizontalLayoutSKUFilename)

        self.buttonSubmit = QPushButton("Generate")
        self.buttonSubmit.clicked.connect(self.onButtonSubmitClick)
        self.verticalLayout.addWidget(self.buttonSubmit)
        self.setLayout(self.verticalLayout)


    def onSKUReadMethodToggled(self):
        global SKU_READ_METHOD, EXPORT_METHOD
        radiobuttonReadSKU = self.sender()
        if radiobuttonReadSKU.method in ['readFromLocal', 'readFromGoogleSpreadSheet']:
            SKU_READ_METHOD = radiobuttonReadSKU.method
            if radiobuttonReadSKU.isChecked():
                if radiobuttonReadSKU.method == 'readFromLocal':
                    self.labelSkuIdFilename.setHidden(False)
                    self.lineEditSkuIdFilename.show()
                    self.buttonSKUFilePicker.show()
                    self.labelSkuReadMethod.hide()
                    self.radiobuttonExportUrlLocal.hide()
                    self.radiobuttonExportUrlGoogleSpreadSheet.hide()
                else:
                    self.labelSkuIdFilename.setHidden(True)
                    self.lineEditSkuIdFilename.hide()
                    self.buttonSKUFilePicker.hide()
                    self.labelSkuReadMethod.show()
                    self.radiobuttonExportUrlLocal.show()
                    self.radiobuttonExportUrlGoogleSpreadSheet.show()

    def onExportMethodToggled(self):
        global SKU_READ_METHOD, EXPORT_METHOD
        radiobuttonExportUrl = self.sender()
        if radiobuttonExportUrl.method in ['exportToLocal', 'exportToGoogleSpreadSheet']:
            SKU_READ_METHOD = 'readFromGoogleSpreadSheet'
            EXPORT_METHOD = radiobuttonExportUrl.method

    def onButtonSKUFilePickerClick(self):
        filepath = self.filePicker()
        if (filepath):
            SKU_FILE = filepath
            self.lineEditSkuIdFilename.setText(SKU_FILE)

    def onButtonImageFolderPickerClick(self):
        filepath = self.folderPicker()
        if (filepath):
            IMAGES_FOLDER = filepath
            self.lineEditImageFolder.setText(IMAGES_FOLDER)

    def onButtonSubmitClick(self):
        global BASE_IMAGE_PATH_URL, SKU_ID_EXTRAS, IMAGES_FOLDER, SKU_FILE, SKU_READ_METHOD
        msg = ""
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

        if (len(self.lineEditSkuIdFilename.text()) > 0):
            SKU_FILE = self.lineEditSkuIdFilename.text()
        else:
            msg = "No SKU filename provided"

        SKU_ID_EXTRAS = list(set(self.lineEditSkuIdVariations.text().strip().split(',')))
        SKU_ID_EXTRAS = list(map(lambda x: (x.strip()), SKU_ID_EXTRAS))

        if (len(msg) > 0):
            self.showDialog(msg)
            return

        result = generateURLList()
        self.showDialog(result[1], result[0])

    def showDialog(self, msgText = "Something went wrong!", status = False):
        msg = QMessageBox()
        msg.setWindowTitle("Success" if status else "Oops!!!")
        msg.setIcon(QMessageBox.Information if status else QMessageBox.Critical)
        msg.setText("Process completed" if status else "Process failed")
        msg.setDetailedText(f"The details are as follows:\n{msgText}")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.exec_()

    def folderPicker(self):
        directory = QFileDialog.getExistingDirectory(self, "Select a folder", "./")
        if (len(directory) > 0):
            return directory
        else:
            return False

    def filePicker(self):
        filepath = QFileDialog.getOpenFileName(self, "Select a file", "./", "All Files (*);;Text Files (*.txt);;CSV Files (*.csv);;") #
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
    SKU_READ_METHOD = config['SKU_READ_METHOD']
    EXPORT_METHOD = config['EXPORT_METHOD']
    SAMPLE_SPREADSHEET_ID = config['SAMPLE_SPREADSHEET_ID']
    READ_RANGE_NAME = config['READ_RANGE_NAME']
    WRITE_RANGE_NAME_1 = config['WRITE_RANGE_NAME_1']
    WRITE_RANGE_NAME_2 = config['WRITE_RANGE_NAME_2']
    window()
