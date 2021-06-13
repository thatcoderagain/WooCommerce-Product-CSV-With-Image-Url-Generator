import os, sys, json, csv
import pandas
import gspread
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from google.oauth2 import service_account
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

SERVICE_ACCOUNT_FILE = './docs-316004-54c2dd979ce3.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDENTIALS = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
SPREADSHEET_ID = '1S0vxL_-7bGZ64I4b3s86L4T0G8PNf6qoRxr4YxwUsd4'
SKU_READ_RANGE = 'Data!C2:C1000'

BASE_IMAGE_PATH_URL = 'https://homefabindia.com/wp-content/uploads/images/products/curtains/'
SKU_ID_EXTRAS = ['5F','6F','7F','8F','9F','Setof2']
IMAGES_FOLDER = 'Converted images'
SKU_FILE = 'SKU_List.txt'
EXCEL_FILE = 'Products.xlsx'
SHEET_NAME = 'Sheet1'
SKU_READ_METHOD = 'readFromLocal'
EXPORT_METHOD = 'exportToLocal'
PRODUCT_TITLE = ''
PRIMARY_PRODUCT_SKU = ''
PRODUCT_FIXED_IMAGES_LINKS = ''
VARIABLE_POINTS_START_CELL = ''
VARIABLE_POINTS_END_CELL = ''
FIXED_POINTS_START_CELL = ''
FIXED_POINTS_END_CELL = ''

SUCCESS = 1
WARNING = 2
ERROR = 3
S_MESSAGE = ''
W_MESSAGE = ''
E_MESSAGE = ''

def createConfig():
    configDict = {}
    configDict['SPREADSHEET_ID'] = SPREADSHEET_ID
    configDict['SKU_READ_RANGE'] = SKU_READ_RANGE
    configDict['BASE_IMAGE_PATH_URL'] = BASE_IMAGE_PATH_URL
    configDict['SKU_ID_EXTRAS'] = sorted(SKU_ID_EXTRAS)
    configDict['IMAGES_FOLDER'] = IMAGES_FOLDER
    configDict['SKU_FILE'] = SKU_FILE
    configDict['EXCEL_FILE'] = EXCEL_FILE
    configDict['SHEET_NAME'] = SHEET_NAME
    configDict['SKU_READ_METHOD'] = SKU_READ_METHOD
    configDict['EXPORT_METHOD'] = EXPORT_METHOD
    configDict['PRODUCT_TITLE'] = PRODUCT_TITLE
    configDict['PRIMARY_PRODUCT_SKU'] = PRIMARY_PRODUCT_SKU
    configDict['PRODUCT_FIXED_IMAGES_LINKS'] = PRODUCT_FIXED_IMAGES_LINKS
    configDict['VARIABLE_POINTS_START_CELL'] = VARIABLE_POINTS_START_CELL
    configDict['VARIABLE_POINTS_END_CELL'] = VARIABLE_POINTS_END_CELL
    configDict['FIXED_POINTS_START_CELL'] = FIXED_POINTS_START_CELL
    configDict['FIXED_POINTS_END_CELL'] = FIXED_POINTS_END_CELL
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
    service = build('sheets', 'v4', credentials=CREDENTIALS)
    value_range_body = {}
    value_range_body['values'] = data
    response = service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=range,
                                                    valueInputOption="USER_ENTERED", body=value_range_body).execute()


def clearSheet(sheetName):
    service = build('sheets', 'v4', credentials=CREDENTIALS)
    rangeAll = '{0}!A1:ZZ'.format(sheetName)
    response = service.spreadsheets().values().clear(spreadsheetId=SPREADSHEET_ID, body={},
                                                    range='{0}!A1:Z'.format(sheetName)).execute()


def readSpreadSheet(read_range):
    service = build('sheets', 'v4', credentials=CREDENTIALS)
    return service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=read_range).execute()


def readSKUIdsFromSpreadSheet():
    response = readSpreadSheet(SKU_READ_RANGE)
    values = response.get('values', [])
    return list(map(lambda x: (x[0]), values))


def copyDataToAnotherSheet(source, target):
    read_range = source+'!A1:ZZ'
    write_range = target+'!A1'
    response = readSpreadSheet(read_range)
    values = response.get('values', [])
    print("Read Response: ", values)
    response = writeToSpreadSheet(write_range, values)
    print("Copied Response: ", response)


def createOrClearSheet(sheet_name):
    client = gspread.authorize(CREDENTIALS)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    worksheet_list = spreadsheet.worksheets()
    worksheet_list = list(map(lambda x: x.title, worksheet_list))
    if (sheet_name in worksheet_list):
        print(f"Deleting old sheet '{sheet_name}'")
        worksheet = spreadsheet.worksheet(sheet_name)
        spreadsheet.del_worksheet(worksheet)
    print(f"Creating New sheet '{sheet_name}'")
    return spreadsheet.add_worksheet(title=sheet_name, rows="10", cols="10")


def writeToCSV(CSV_FILE_NAME, sheet_content):
    output_file = f'./csv/{CSV_FILE_NAME}.csv'
    with open(output_file, 'w') as file:
        writer = csv.writer(file, quoting=csv.QUOTE_NONNUMERIC)
        writer.writerows(sheet_content)
    file.close()


def exportSheetToCSV(sheet_name, csv_file):
    service = build('sheets', 'v4', credentials=CREDENTIALS)
    sheet_range = sheet_name+'!A1:ZZ'
    response = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=sheet_range).execute()
    writeToCSV(csv_file, response.get('values', []))


def customFilter(string, targets):
    for target in targets:
        string = string[len(target):] if string.startswith(target) else string
        string = string[:len(string)-len(target)] if string.endswith(target) else string
    return string


def generateURLList():
    global S_MESSAGE, W_MESSAGE, E_MESSAGE
    updateConfig(createConfig())
    if (os.path.isdir(IMAGES_FOLDER) == False):
        E_MESSAGE = "Oops! Images folder not found"
        return [ERROR, E_MESSAGE]

    listOfFiles = os.listdir(IMAGES_FOLDER)
    images = filter(lambda x: x.endswith('.jpg'), listOfFiles)
    images = list(map(lambda x: BASE_IMAGE_PATH_URL+x, images))
    images = sorted(images)

    if (len(images) == 0):
        E_MESSAGE = "Oops! Images not found in the selected folder"
        return [ERROR, E_MESSAGE]

    if SKU_READ_METHOD == 'readExcelAndExportProductToGoogleSpreadSheet' or EXPORT_METHOD == 'exportUrlToGoogleSpreadSheet':
        variablesData = list()
        variablesData.append(['Product Id', 1])
        variablesData.append(['Low stock amount', 5])
        variablesData.append(['Primary Product Title', PRODUCT_TITLE])
        variablesData.append(['Fixed Bullet Point Start Column', FIXED_POINTS_START_CELL])
        variablesData.append(['Fixed Bullet Point End Column', FIXED_POINTS_END_CELL])
        variablesData.append(['Variable Bullet Point Start Column', VARIABLE_POINTS_START_CELL])
        variablesData.append(['Variable Bullet Point End Column', VARIABLE_POINTS_END_CELL])
        variablesData.append(['Primary Image SKU ID', PRIMARY_PRODUCT_SKU])
        variablesData.append(['Static Image Links for Primary Image', PRODUCT_FIXED_IMAGES_LINKS])
        response = clearSheet('Variables')
        response = writeToSpreadSheet('Variables!A1', variablesData)
        print("Read Response: ", response)

    if SKU_READ_METHOD == 'readFromLocal':
        if (os.path.exists(SKU_FILE) == False):
            E_MESSAGE = "Oops! SKU file not found"
            return [ERROR, E_MESSAGE]
        skuFile = open(SKU_FILE, 'r')
        skuIds = skuFile.readlines()
    elif SKU_READ_METHOD == 'readFromGoogleSpreadSheet':
        skuIds = readSKUIdsFromSpreadSheet()
    elif SKU_READ_METHOD == 'readExcelAndExportProductToGoogleSpreadSheet':
        if (os.path.exists(EXCEL_FILE) == False):
            E_MESSAGE = "Oops! Product excel file not found"
            return [ERROR, E_MESSAGE]
        data = readExcel(EXCEL_FILE, SHEET_NAME)
        print("Excel Data: \n", data)
        response = clearSheet('Data')
        print("Read Response: ", response)
        response = writeToSpreadSheet('Data!A1', data)
        print("Read Response: ", response)
        skuIds = readSKUIdsFromSpreadSheet()

    skuIds = map(lambda x: (x.strip()), skuIds)
    skuIds = list(map(lambda x: (customFilter(x, SKU_ID_EXTRAS)), skuIds))

    if (len(skuIds) == 0):
        E_MESSAGE = "Oops! No SKU ids found for the products"
        return [ERROR, E_MESSAGE]

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
                imagesSet[skuId] = list(sorted(set(imagesSet[skuId])))

    if SKU_READ_METHOD == 'readFromLocal' or EXPORT_METHOD == 'exportToLocal':
        AD3 = open('./imagesUrls/AD3.csv', 'w')
        AN3 = open('./imagesUrls/AN3.csv', 'w')
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

        primarySku = customFilter(PRIMARY_PRODUCT_SKU, SKU_ID_EXTRAS)
        if primarySku in skuIds:
            primaryImageLinks = ",".join(imagesSet[primarySku])
        else:
            primaryImageLinks = ""
            W_MESSAGE = f"Primary SKU '{primarySku}' doesn't match with any product variant so unable to generate URLs for the primary product"

        if (len(PRODUCT_FIXED_IMAGES_LINKS) > 0):
            fixedLinks = PRODUCT_FIXED_IMAGES_LINKS.split(',')
            fixedLinks = list(map(lambda x: x.strip(), fixedLinks))
            fixedLinks = list(filter(lambda x: x.endswith('.jpg'), fixedLinks))
            fixedLinks = ",".join(fixedLinks)
        else:
            fixedLinks = ""

        if len(primaryImageLinks) > 0 and len(fixedLinks) > 0:
            primaryImageLinks = primaryImageLinks+','+fixedLinks
        elif len(fixedLinks) > 0:
            primaryImageLinks = fixedLinks
        primaryImageLinks = primaryImageLinks
        print("primaryImageLinks: ", primaryImageLinks)

        AD2 = [[primaryImageLinks]]+AD3
        response= writeToSpreadSheet('Compiled!AD2', AD2)
        print("Read Response: ", response)
        response = writeToSpreadSheet('Compiled!AN3', AN3)
        print("Read Response: ", response)

        productSheetName = EXCEL_FILE[EXCEL_FILE.rindex('/')+1 if '/' in EXCEL_FILE else 0:EXCEL_FILE.rindex('.')]
        createOrClearSheet(productSheetName)
        copyDataToAnotherSheet('Compiled', productSheetName)
        exportSheetToCSV(productSheetName, productSheetName)

    updateConfig(createConfig())
    S_MESSAGE = "All done Khushi Goyal :*"
    if (len(W_MESSAGE) > 0):
        return [WARNING, W_MESSAGE]
    if (len(E_MESSAGE) > 0):
        return [ERROR, E_MESSAGE]
    return [SUCCESS, S_MESSAGE]


class Widgets(QWidget):
    def __init__(self, **kwargs):
        super(Widgets, self).__init__()

        self.setWindowTitle("Khushi Tool")
        self.setGeometry(100,100,900,250)
        self.move(200,200)

        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.setLayout(self.verticalLayout)

        self.labelBaseImageUrl = QLabel()
        self.labelBaseImageUrl.setText("Base image url ")
        self.lineEditBaseImageUrl = QLineEdit()
        self.lineEditBaseImageUrl.setFixedWidth(712)
        self.lineEditBaseImageUrl.setText(BASE_IMAGE_PATH_URL)
        self.horizontalLayoutBaseImageUrl = QHBoxLayout()
        self.horizontalLayoutBaseImageUrl.addWidget(self.labelBaseImageUrl)
        self.horizontalLayoutBaseImageUrl.addWidget(self.lineEditBaseImageUrl)
        self.verticalLayout.addLayout(self.horizontalLayoutBaseImageUrl)

        self.labelSkuIdVariations = QLabel()
        self.labelSkuIdVariations.setText("SKU variations ")
        self.lineEditSkuIdVariations = QLineEdit()
        self.lineEditSkuIdVariations.setFixedWidth(712)
        self.lineEditSkuIdVariations.setText(",".join(SKU_ID_EXTRAS))
        self.horizontalLayoutSKUVariations = QHBoxLayout()
        self.horizontalLayoutSKUVariations.addWidget(self.labelSkuIdVariations)
        self.horizontalLayoutSKUVariations.addWidget(self.lineEditSkuIdVariations)
        self.verticalLayout.addLayout(self.horizontalLayoutSKUVariations)

        self.labelImagesFolder = QLabel()
        self.labelImagesFolder.setText("Images Folder ")
        self.lineEditImageFolder = QLineEdit()
        self.lineEditImageFolder.setFixedWidth(500)
        self.lineEditImageFolder.setText(IMAGES_FOLDER)
        self.buttonSkuFilePicker = QPushButton("Browse (Images) ")
        self.buttonSkuFilePicker.setFixedWidth(205)
        self.buttonSkuFilePicker.clicked.connect(self.onButtonImageFolderPickerClick)
        self.horizontalLayoutImagesFolderName = QHBoxLayout()
        self.horizontalLayoutImagesFolderName.addWidget(self.labelImagesFolder)
        self.horizontalLayoutImagesFolderName.addWidget(self.lineEditImageFolder)
        self.horizontalLayoutImagesFolderName.addWidget(self.buttonSkuFilePicker)
        self.verticalLayout.addLayout(self.horizontalLayoutImagesFolderName)

        self.labelSkuReadMethod = QLabel()
        self.labelSkuReadMethod.setText("SKU Read Method")
        self.radiobuttonReadSkuLocal = QRadioButton("Local File")
        self.radiobuttonReadSkuLocal.setFixedWidth(170)
        self.radiobuttonReadSkuLocal.method = "readFromLocal"
        self.radiobuttonReadSkuLocal.toggled.connect(self.onSKUReadMethodToggled)
        self.radiobuttonReadSkuGoogleSpreadSheet = QRadioButton("Google SpreadSheet")
        self.radiobuttonReadSkuGoogleSpreadSheet.setFixedWidth(200)
        self.radiobuttonReadSkuGoogleSpreadSheet.method = "readFromGoogleSpreadSheet"
        self.radiobuttonReadSkuGoogleSpreadSheet.toggled.connect(self.onSKUReadMethodToggled)
        self.radiobuttonUploadProductToGoogleSpreadSheet = QRadioButton("Import Product + Upload to SpreadSheet")
        self.radiobuttonUploadProductToGoogleSpreadSheet.setFixedWidth(330)
        self.radiobuttonUploadProductToGoogleSpreadSheet.method = "readExcelAndExportProductToGoogleSpreadSheet"
        self.radiobuttonUploadProductToGoogleSpreadSheet.toggled.connect(self.onSKUReadMethodToggled)
        self.buttonGroupSkuRead = QButtonGroup()
        self.buttonGroupSkuRead.addButton(self.radiobuttonReadSkuLocal)
        self.buttonGroupSkuRead.addButton(self.radiobuttonReadSkuGoogleSpreadSheet)
        self.buttonGroupSkuRead.addButton(self.radiobuttonUploadProductToGoogleSpreadSheet)
        self.horizontalLayoutSkuReadMethod = QHBoxLayout()
        self.horizontalLayoutSkuReadMethod.addWidget(self.labelSkuReadMethod)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonReadSkuLocal)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonReadSkuGoogleSpreadSheet)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonUploadProductToGoogleSpreadSheet)
        self.verticalLayout.addLayout(self.horizontalLayoutSkuReadMethod)


        self.labelSkuReadMethod = QLabel()
        self.labelSkuReadMethod.setText("Export Method")
        self.radiobuttonExportUrlToLocal = QRadioButton("Generate .csv files")
        self.radiobuttonExportUrlToLocal.method = "exportToLocal"
        self.radiobuttonExportUrlToLocal.setFixedWidth(170)
        self.radiobuttonExportUrlToLocal.toggled.connect(self.onExportMethodToggled)
        self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet = QRadioButton("Export Urls to SpreadSheet")
        self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet.setFixedWidth(535)
        self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet.method = "exportUrlToGoogleSpreadSheet"
        self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet.toggled.connect(self.onExportMethodToggled)
        self.buttonGroupExportUrl = QButtonGroup()
        self.buttonGroupExportUrl.addButton(self.radiobuttonExportUrlToLocal)
        self.buttonGroupExportUrl.addButton(self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet)
        self.horizontalLayoutSkuReadMethod = QHBoxLayout()
        self.horizontalLayoutSkuReadMethod.addWidget(self.labelSkuReadMethod)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonExportUrlToLocal)
        self.horizontalLayoutSkuReadMethod.addWidget(self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet)
        self.verticalLayout.addLayout(self.horizontalLayoutSkuReadMethod)

        self.labelSkuIdFilename = QLabel()
        self.labelSkuIdFilename.setText("SKU Filename ")
        self.lineEditSkuIdFilename = QLineEdit()
        self.lineEditSkuIdFilename.setFixedWidth(500)
        self.lineEditSkuIdFilename.setText(SKU_FILE)
        self.buttonSkuFilePicker = QPushButton("Browse (SKU List.txt) ")
        self.buttonSkuFilePicker.setFixedWidth(205)
        self.horizontalLayoutSKUFilename = QHBoxLayout()
        self.buttonSkuFilePicker.clicked.connect(self.onButtonSkuFilePickerClick)
        self.horizontalLayoutSKUFilename.addWidget(self.labelSkuIdFilename)
        self.horizontalLayoutSKUFilename.addWidget(self.lineEditSkuIdFilename)
        self.horizontalLayoutSKUFilename.addWidget(self.buttonSkuFilePicker)
        self.verticalLayout.addLayout(self.horizontalLayoutSKUFilename)

        self.labelExcelFilename = QLabel()
        self.labelExcelFilename.setText("Product File ")
        self.lineEditExcelFilename = QLineEdit()
        self.lineEditExcelFilename.setFixedWidth(500)
        self.lineEditExcelFilename.setText(EXCEL_FILE)
        self.buttonExcelFilename = QPushButton("Browse (Products.xlsx)")
        self.buttonExcelFilename.setFixedWidth(205)
        self.horizontalLayoutExcelFilePicker = QHBoxLayout()
        self.horizontalLayoutExcelFilePicker.addWidget(self.labelExcelFilename)
        self.horizontalLayoutExcelFilePicker.addWidget(self.lineEditExcelFilename)
        self.buttonExcelFilename.clicked.connect(self.onButtonExcelFilePickerClick)
        self.horizontalLayoutExcelFilePicker.addWidget(self.buttonExcelFilename)
        self.verticalLayout.addLayout(self.horizontalLayoutExcelFilePicker)

        self.labelPrimaryProductTitle = QLabel()
        self.labelPrimaryProductTitle.setText("Product Title ");
        self.lineEditPrimaryProductTitle = QLineEdit()
        self.lineEditPrimaryProductTitle.setFixedWidth(712)
        self.lineEditPrimaryProductTitle.setText(PRODUCT_TITLE)
        self.horizontalLayoutPrimaryProductTitle = QHBoxLayout()
        self.horizontalLayoutPrimaryProductTitle.addWidget(self.labelPrimaryProductTitle)
        self.horizontalLayoutPrimaryProductTitle.addWidget(self.lineEditPrimaryProductTitle)
        self.verticalLayout.addLayout(self.horizontalLayoutPrimaryProductTitle)

        self.labelFixedBulletPointSpace = QLabel()
        self.labelFixedBulletPointSpace.setFixedWidth(200)
        self.labelFixedBulletPointSpace.setText("Fixed Points")
        self.labelFixedBulletPointStart = QLabel()
        self.labelFixedBulletPointStart.setFixedWidth(200)
        self.labelFixedBulletPointStart.setText("Start Cell ")
        self.lineEditFixedBulletPointStart = QLineEdit()
        self.lineEditFixedBulletPointStart.setFixedWidth(100)
        self.lineEditFixedBulletPointStart.setText(FIXED_POINTS_START_CELL)
        self.labelFixedBulletPointEnd = QLabel()
        self.labelFixedBulletPointEnd.setFixedWidth(200)
        self.labelFixedBulletPointEnd.setText("End Cell ")
        self.lineEditFixedBulletPointEnd = QLineEdit()
        self.lineEditFixedBulletPointEnd.setFixedWidth(100)
        self.lineEditFixedBulletPointEnd.setText(FIXED_POINTS_END_CELL)
        self.horizontalLayoutFixedBulletPoint = QHBoxLayout()
        self.horizontalLayoutFixedBulletPoint.addWidget(self.labelFixedBulletPointSpace, alignment=Qt.AlignLeft | Qt.AlignBottom)
        self.horizontalLayoutFixedBulletPoint.addWidget(self.labelFixedBulletPointStart, alignment=Qt.AlignRight | Qt.AlignBottom)
        self.horizontalLayoutFixedBulletPoint.addWidget(self.lineEditFixedBulletPointStart, alignment=Qt.AlignLeft | Qt.AlignBottom)
        self.horizontalLayoutFixedBulletPoint.addWidget(self.labelFixedBulletPointEnd, alignment=Qt.AlignRight | Qt.AlignBottom)
        self.horizontalLayoutFixedBulletPoint.addWidget(self.lineEditFixedBulletPointEnd, alignment=Qt.AlignRight | Qt.AlignBottom)
        self.verticalLayout.addLayout(self.horizontalLayoutFixedBulletPoint)

        self.labelVariableBulletPointSpace = QLabel()
        self.labelVariableBulletPointSpace.setFixedWidth(200)
        self.labelVariableBulletPointSpace.setText("Variable Points")
        self.labelVariableBulletPointStart = QLabel()
        self.labelVariableBulletPointStart.setFixedWidth(200)
        self.labelVariableBulletPointStart.setText("Start Cell ")
        self.lineEditVariableBulletPointStart = QLineEdit()
        self.lineEditVariableBulletPointStart.setFixedWidth(100)
        self.lineEditVariableBulletPointStart.setText(VARIABLE_POINTS_START_CELL)
        self.labelVariableBulletPointEnd = QLabel()
        self.labelVariableBulletPointEnd.setFixedWidth(200)
        self.labelVariableBulletPointEnd.setText("End Cell ")
        self.lineEditVariableBulletPointEnd = QLineEdit()
        self.lineEditVariableBulletPointEnd.setFixedWidth(100)
        self.lineEditVariableBulletPointEnd.setText(VARIABLE_POINTS_END_CELL)
        self.horizontalLayoutVariableBulletPoint = QHBoxLayout()
        self.horizontalLayoutVariableBulletPoint.addWidget(self.labelVariableBulletPointSpace, alignment=Qt.AlignLeft | Qt.AlignBottom)
        self.horizontalLayoutVariableBulletPoint.addWidget(self.labelVariableBulletPointStart, alignment=Qt.AlignLeft | Qt.AlignBottom)
        self.horizontalLayoutVariableBulletPoint.addWidget(self.lineEditVariableBulletPointStart, alignment=Qt.AlignLeft | Qt.AlignBottom)
        self.horizontalLayoutVariableBulletPoint.addWidget(self.labelVariableBulletPointEnd, alignment=Qt.AlignRight | Qt.AlignBottom)
        self.horizontalLayoutVariableBulletPoint.addWidget(self.lineEditVariableBulletPointEnd, alignment=Qt.AlignRight | Qt.AlignBottom)
        self.verticalLayout.addLayout(self.horizontalLayoutVariableBulletPoint)

        self.labelPrimaryProductSkuId = QLabel()
        self.labelPrimaryProductSkuId.setText("Primary Product SKU ");
        self.lineEditPrimaryProductSkuId = QLineEdit()
        self.lineEditPrimaryProductSkuId.setFixedWidth(712)
        self.lineEditPrimaryProductSkuId.setText(PRIMARY_PRODUCT_SKU)
        self.horizontalLayoutPrimaryProductSkuId = QHBoxLayout()
        self.horizontalLayoutPrimaryProductSkuId.addWidget(self.labelPrimaryProductSkuId)
        self.horizontalLayoutPrimaryProductSkuId.addWidget(self.lineEditPrimaryProductSkuId)
        self.verticalLayout.addLayout(self.horizontalLayoutPrimaryProductSkuId)

        self.labelPrimaryProductFixedImageLinks = QLabel()
        self.labelPrimaryProductFixedImageLinks.setText("Product Fixed Image Links ");
        self.lineEditPrimaryProductFixedImageLinks = QLineEdit()
        self.lineEditPrimaryProductFixedImageLinks.setFixedWidth(712)
        self.lineEditPrimaryProductFixedImageLinks.setText(PRODUCT_FIXED_IMAGES_LINKS)
        self.horizontalLayoutPrimaryProductFixedImageLinks = QHBoxLayout()
        self.horizontalLayoutPrimaryProductFixedImageLinks.addWidget(self.labelPrimaryProductFixedImageLinks)
        self.horizontalLayoutPrimaryProductFixedImageLinks.addWidget(self.lineEditPrimaryProductFixedImageLinks)
        self.verticalLayout.addLayout(self.horizontalLayoutPrimaryProductFixedImageLinks)

        self.buttonSubmit = QPushButton("Generate")
        self.buttonSubmit.clicked.connect(self.onButtonSubmitClick)
        self.verticalLayout.addWidget(self.buttonSubmit)

        if SKU_READ_METHOD == 'readFromLocal':
            self.radiobuttonReadSkuLocal.setChecked(True)
        elif SKU_READ_METHOD == 'readFromGoogleSpreadSheet':
            self.radiobuttonReadSkuGoogleSpreadSheet.setChecked(True)
            if EXPORT_METHOD == 'exportToLocal':
                self.radiobuttonExportUrlToLocal.setChecked(True)
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
                    self.radiobuttonExportUrlToLocal.hide()
                    self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet.hide()
                    self.labelExcelFilename.setHidden(True)
                    self.lineEditExcelFilename.hide()
                    self.buttonExcelFilename.hide()
                    self.labelPrimaryProductTitle.setHidden(True)
                    self.lineEditPrimaryProductTitle.hide()
                    self.labelFixedBulletPointSpace.setHidden(True)
                    self.labelFixedBulletPointStart.setHidden(True)
                    self.lineEditFixedBulletPointStart.hide()
                    self.labelFixedBulletPointEnd.setHidden(True)
                    self.lineEditFixedBulletPointEnd.hide()
                    self.labelVariableBulletPointSpace.setHidden(True)
                    self.labelVariableBulletPointStart.setHidden(True)
                    self.lineEditVariableBulletPointStart.hide()
                    self.labelVariableBulletPointEnd.setHidden(True)
                    self.lineEditVariableBulletPointEnd.hide()
                    self.labelPrimaryProductSkuId.setHidden(True)
                    self.lineEditPrimaryProductSkuId.hide()
                    self.labelPrimaryProductFixedImageLinks.setHidden(True)
                    self.lineEditPrimaryProductFixedImageLinks.hide()
                elif radiobuttonReadSKU.method == 'readFromGoogleSpreadSheet':
                    self.labelSkuIdFilename.setHidden(True)
                    self.lineEditSkuIdFilename.hide()
                    self.buttonSkuFilePicker.hide()
                    self.labelSkuReadMethod.show()
                    self.radiobuttonExportUrlToLocal.show()
                    self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet.show()
                    self.labelExcelFilename.setHidden(True)
                    self.lineEditExcelFilename.hide()
                    self.buttonExcelFilename.hide()
                    self.radiobuttonExportUrlToLocal.setChecked(True)
                elif radiobuttonReadSKU.method == 'readExcelAndExportProductToGoogleSpreadSheet':
                    self.labelSkuIdFilename.setHidden(True)
                    self.lineEditSkuIdFilename.hide()
                    self.buttonSkuFilePicker.hide()
                    self.labelSkuReadMethod.hide()
                    self.radiobuttonExportUrlToLocal.hide()
                    self.radiobuttonReadExcelAndExportUrlToGoogleSpreadSheet.hide()
                    self.labelExcelFilename.setHidden(False)
                    self.lineEditExcelFilename.show()
                    self.buttonExcelFilename.show()
                    self.labelPrimaryProductTitle.setHidden(False)
                    self.lineEditPrimaryProductTitle.show()
                    self.labelFixedBulletPointSpace.setHidden(False)
                    self.labelFixedBulletPointStart.setHidden(False)
                    self.lineEditFixedBulletPointStart.show()
                    self.labelFixedBulletPointEnd.setHidden(False)
                    self.lineEditFixedBulletPointEnd.show()
                    self.labelVariableBulletPointSpace.setHidden(False)
                    self.labelVariableBulletPointStart.setHidden(False)
                    self.lineEditVariableBulletPointStart.show()
                    self.labelVariableBulletPointEnd.setHidden(False)
                    self.lineEditVariableBulletPointEnd.show()
                    self.labelPrimaryProductSkuId.setHidden(False)
                    self.lineEditPrimaryProductSkuId.show()
                    self.labelPrimaryProductFixedImageLinks.setHidden(False)
                    self.lineEditPrimaryProductFixedImageLinks.show()

    def onExportMethodToggled(self):
        global SKU_READ_METHOD, EXPORT_METHOD
        radiobuttonExportUrl = self.sender()
        if radiobuttonExportUrl.method in ['exportToLocal', 'exportUrlToGoogleSpreadSheet']:
            EXPORT_METHOD = radiobuttonExportUrl.method
            if EXPORT_METHOD == 'exportUrlToGoogleSpreadSheet':
                self.labelPrimaryProductTitle.setHidden(False)
                self.lineEditPrimaryProductTitle.show()
                self.labelFixedBulletPointSpace.setHidden(False)
                self.labelFixedBulletPointStart.setHidden(False)
                self.lineEditFixedBulletPointStart.show()
                self.labelFixedBulletPointEnd.setHidden(False)
                self.lineEditFixedBulletPointEnd.show()
                self.labelVariableBulletPointSpace.setHidden(False)
                self.labelVariableBulletPointStart.setHidden(False)
                self.lineEditVariableBulletPointStart.show()
                self.labelVariableBulletPointEnd.setHidden(False)
                self.lineEditVariableBulletPointEnd.show()
                self.labelPrimaryProductSkuId.setHidden(False)
                self.lineEditPrimaryProductSkuId.show()
                self.labelPrimaryProductFixedImageLinks.setHidden(False)
                self.lineEditPrimaryProductFixedImageLinks.show()
            else:
                self.labelPrimaryProductTitle.setHidden(True)
                self.lineEditPrimaryProductTitle.hide()
                self.labelFixedBulletPointSpace.setHidden(True)
                self.labelFixedBulletPointStart.setHidden(True)
                self.lineEditFixedBulletPointStart.hide()
                self.labelFixedBulletPointEnd.setHidden(True)
                self.lineEditFixedBulletPointEnd.hide()
                self.labelVariableBulletPointSpace.setHidden(True)
                self.labelVariableBulletPointStart.setHidden(True)
                self.lineEditVariableBulletPointStart.hide()
                self.labelVariableBulletPointEnd.setHidden(True)
                self.lineEditVariableBulletPointEnd.hide()
                self.labelPrimaryProductSkuId.setHidden(True)
                self.lineEditPrimaryProductSkuId.hide()
                self.labelPrimaryProductFixedImageLinks.setHidden(True)
                self.lineEditPrimaryProductFixedImageLinks.hide()

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
        global BASE_IMAGE_PATH_URL, SKU_ID_EXTRAS, IMAGES_FOLDER, SKU_FILE, SKU_READ_METHOD, EXPORT_METHOD, EXCEL_FILE, PRODUCT_TITLE, FIXED_POINTS_START_CELL, FIXED_POINTS_END_CELL, VARIABLE_POINTS_START_CELL, VARIABLE_POINTS_END_CELL, PRIMARY_PRODUCT_SKU, PRODUCT_FIXED_IMAGES_LINKS
        E_MESSAGE = ""
        self.buttonSubmit.setEnabled(False)
        if (self.lineEditBaseImageUrl.text().strip().startswith('https://')):
            BASE_IMAGE_PATH_URL = self.lineEditBaseImageUrl.text().strip()
            if (BASE_IMAGE_PATH_URL.endswith('/') == False):
                BASE_IMAGE_PATH_URL = BASE_IMAGE_PATH_URL+'/'
        else:
            E_MESSAGE = "Invalid base URL provided"

        if (len(self.lineEditImageFolder.text().strip()) > 0):
            IMAGES_FOLDER = self.lineEditImageFolder.text().strip()
        else:
            E_MESSAGE = "No images folder provided"

        SKU_ID_EXTRAS = list(map(lambda x: (x.strip()), list(set(self.lineEditSkuIdVariations.text().strip().split(',')))))

        if (SKU_READ_METHOD == 'readFromLocal'):
            if (len(self.lineEditSkuIdFilename.text().strip()) > 0):
                SKU_FILE = self.lineEditSkuIdFilename.text().strip()
            else:
                E_MESSAGE = "No SKU filename provided"
        elif (SKU_READ_METHOD == 'readFromGoogleSpreadSheet'):
            pass
        elif (SKU_READ_METHOD == 'readExcelAndExportProductToGoogleSpreadSheet'):
            if (len(self.lineEditExcelFilename.text().strip()) > 0):
                SKU_FILE = self.lineEditExcelFilename.text().strip()
            else:
                E_MESSAGE = "No SKU filename provided"

        if SKU_READ_METHOD == 'readExcelAndExportProductToGoogleSpreadSheet' or EXPORT_METHOD == 'exportUrlToGoogleSpreadSheet':
            if (len(self.lineEditPrimaryProductTitle.text().strip()) > 0):
                PRODUCT_TITLE = self.lineEditPrimaryProductTitle.text().strip()
            else:
                E_MESSAGE = "No product title provided"

            if (len(self.lineEditFixedBulletPointStart.text().strip()) > 0):
                FIXED_POINTS_START_CELL = self.lineEditFixedBulletPointStart.text().strip()
            else:
                E_MESSAGE = "No start cell of fixed bullet point provided"

            if (len(self.lineEditFixedBulletPointEnd.text().strip()) > 0):
                FIXED_POINTS_END_CELL = self.lineEditFixedBulletPointEnd.text().strip()
            else:
                E_MESSAGE = "No end cell of fixed bullet point provided"

            if (len(self.lineEditVariableBulletPointStart.text().strip()) > 0):
                VARIABLE_POINTS_START_CELL = self.lineEditVariableBulletPointStart.text().strip()
            else:
                E_MESSAGE = "No start cell of variable bullet point provided"

            if (len(self.lineEditVariableBulletPointEnd.text().strip()) > 0):
                VARIABLE_POINTS_END_CELL = self.lineEditVariableBulletPointEnd.text().strip()
            else:
                E_MESSAGE = "No end cell of variable bullet points provided"

            if (len(self.lineEditPrimaryProductSkuId.text().strip()) > 0):
                PRIMARY_PRODUCT_SKU = self.lineEditPrimaryProductSkuId.text().strip()
            else:
                E_MESSAGE = "No SKU for primary product provided"

            PRODUCT_FIXED_IMAGES_LINKS = self.lineEditPrimaryProductFixedImageLinks.text().strip()

        if (len(E_MESSAGE) > 0):
            self.showDialog(E_MESSAGE, ERROR)
            return

        result = generateURLList()
        self.showDialog(result[1], result[0])
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

    def showDialog(self, msgText, status = ERROR):
        global S_MESSAGE, W_MESSAGE, E_MESSAGE
        msg = QMessageBox()
        if status == SUCCESS:
            msg.setWindowTitle("Success")
            msg.setIcon(QMessageBox.Information)
            msg.setText(msgText)
        elif status == WARNING:
            msg.setWindowTitle("Warning")
            msg.setIcon(QMessageBox.Warning)
            msg.setText(msgText)
        elif status == ERROR:
            msg.setWindowTitle("ERROR")
            msg.setIcon(QMessageBox.Critical)
            msg.setText(msgText)

        msg.setDetailedText(f"The details are as follows:\n{msgText}")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.exec_()
        self.buttonSubmit.setEnabled(True)
        S_MESSAGE = W_MESSAGE = E_MESSAGE = ''


def window():
    app = QApplication([])

    widget = Widgets()
    widget.show()

    app.exec_()


if __name__ == '__main__':
    config = config()
    SPREADSHEET_ID = config['SPREADSHEET_ID'] if 'SPREADSHEET_ID' in config else SPREADSHEET_ID
    SKU_READ_RANGE = config['SKU_READ_RANGE'] if 'SKU_READ_RANGE' in config else SKU_READ_RANGE
    BASE_IMAGE_PATH_URL = config['BASE_IMAGE_PATH_URL'] if 'BASE_IMAGE_PATH_URL' in config else BASE_IMAGE_PATH_URL
    SKU_ID_EXTRAS = config['SKU_ID_EXTRAS'] if 'SKU_ID_EXTRAS' in config else SKU_ID_EXTRAS
    IMAGES_FOLDER = config['IMAGES_FOLDER'] if 'IMAGES_FOLDER' in config else IMAGES_FOLDER
    SKU_FILE = config['SKU_FILE'] if 'SKU_FILE' in config else SKU_FILE
    EXCEL_FILE = config['EXCEL_FILE'] if 'EXCEL_FILE' in config else EXCEL_FILE
    SHEET_NAME = config['SHEET_NAME'] if 'SHEET_NAME' in config else SHEET_NAME
    SKU_READ_METHOD = config['SKU_READ_METHOD'] if 'SKU_READ_METHOD' in config else SKU_READ_METHOD
    EXPORT_METHOD = config['EXPORT_METHOD'] if 'EXPORT_METHOD' in config else EXPORT_METHOD
    PRODUCT_TITLE = config['PRODUCT_TITLE'] if 'PRODUCT_TITLE' in config else PRODUCT_TITLE
    PRIMARY_PRODUCT_SKU = config['PRIMARY_PRODUCT_SKU'] if 'PRIMARY_PRODUCT_SKU' in config else PRIMARY_PRODUCT_SKU
    PRODUCT_FIXED_IMAGES_LINKS = config['PRODUCT_FIXED_IMAGES_LINKS'] if 'PRODUCT_FIXED_IMAGES_LINKS' in config else PRODUCT_FIXED_IMAGES_LINKS
    VARIABLE_POINTS_START_CELL = config['VARIABLE_POINTS_START_CELL'] if 'VARIABLE_POINTS_START_CELL' in config else VARIABLE_POINTS_START_CELL
    VARIABLE_POINTS_END_CELL = config['VARIABLE_POINTS_END_CELL'] if 'VARIABLE_POINTS_END_CELL' in config else VARIABLE_POINTS_END_CELL
    FIXED_POINTS_START_CELL = config['FIXED_POINTS_START_CELL'] if 'FIXED_POINTS_START_CELL' in config else FIXED_POINTS_START_CELL
    FIXED_POINTS_END_CELL = config['FIXED_POINTS_END_CELL'] if 'FIXED_POINTS_END_CELL' in config else FIXED_POINTS_END_CELL
    window()
