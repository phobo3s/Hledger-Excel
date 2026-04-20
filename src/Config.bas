Attribute VB_Name = "Config"
Option Explicit

' ============================================
' SABITLER — declarations section'da olmali
' ============================================
Public Const HLEDGER_UI_PATH As String = "hledger-ui"
Public Const HLEDGER_UI_ARGS As String = "-w -3 -X TRY --infer-market-prices"

Public Const LOGGING_ENABLED As Boolean = True
Public Const LOG_LEVEL As String = "DEBUG"
Public Const LOG_TO_SHEET As Boolean = True
Public Const LOG_TO_IMMEDIATE As Boolean = True

Public Const COLOR_LIGHT_GREEN As Long = 11854022
Public Const COLOR_DARK_GREEN As Long = 3329434

' ============================================
' DOSYA YOLLARI — ThisWorkbook.Path bazli
' ============================================
Public Property Get DATA_FOLDER() As String
    DATA_FOLDER = ThisWorkbook.path & "\"
End Property

Public Property Get VBA_EXPORT_PATH() As String
    VBA_EXPORT_PATH = ThisWorkbook.path & "\src\"
End Property

Public Property Get TEMP_FILE_ADDR() As String
    TEMP_FILE_ADDR = DATA_FOLDER & "Temp.txt"
End Property

Public Property Get HLEDGER_FILE_ADDR() As String
    HLEDGER_FILE_ADDR = DATA_FOLDER & "Main.hledger"
End Property

Public Property Get COMMODITY_PRICES_FILE() As String
    COMMODITY_PRICES_FILE = DATA_FOLDER & "Commodity-Prices.csv"
End Property

Public Property Get PORTFOLIO_CSV_PATH() As String
    PORTFOLIO_CSV_PATH = DATA_FOLDER & "PortfolioMovements.csv"
End Property

Public Property Get PORTFOLIO_CSV_PATH_SWS() As String
    PORTFOLIO_CSV_PATH_SWS = DATA_FOLDER & "PortfolioMovements_SimplyWallSt.csv"
End Property

Public Property Get PORTFOLIO_CASH_CSV_PATH() As String
    PORTFOLIO_CASH_CSV_PATH = DATA_FOLDER & "PortfolioCashMovements.csv"
End Property

Public Property Get PORTFOLIO_CSV_PATH_INVESTING() As String
    PORTFOLIO_CSV_PATH_INVESTING = DATA_FOLDER & "PortfolioMovements_OpenPositions.csv"
End Property

' ============================================
' HELPER
' ============================================
Public Function ValidateDataFolder() As Boolean
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    ValidateDataFolder = fs.FolderExists(DATA_FOLDER)
End Function






