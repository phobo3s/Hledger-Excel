Attribute VB_Name = "LogManager"
'
' LogManager.bas
' Centralized Logging Framework for Hledger-Excel
' Purpose: Provide structured logging with multiple destinations
'

Option Explicit

' Log Levels
Public Enum LogLevel
    DEBUG_LEVEL = 0
    INFO_LEVEL = 1
    WARNING_LEVEL = 2
    ERROR_LEVEL = 3
End Enum

Private Const LOG_SHEET_NAME As String = "LOGS"

' ============================================
' Public Log Functions
' ============================================

Public Sub LogDebug(msg As String)
    If Config.LOGGING_ENABLED And Config.LOG_LEVEL = "DEBUG" Then
        WriteLog msg, "DEBUG"
    End If
End Sub

Public Sub LogInfo(msg As String)
    If Config.LOGGING_ENABLED Then
        WriteLog msg, "INFO"
    End If
End Sub

Public Sub LogWarning(msg As String)
    If Config.LOGGING_ENABLED Then
        WriteLog msg, "WARNING"
    End If
End Sub

Public Sub LogError(msg As String)
    If Config.LOGGING_ENABLED Then
        WriteLog msg, "ERROR"
    End If
End Sub

' ============================================
' Private Helper: Write to destinations
' ============================================

Private Sub WriteLog(msg As String, level As String)
    Dim timestamp As String
    Dim logEntry As String

    timestamp = Format(Now(), "YYYY-MM-DD HH:MM:SS")
    logEntry = timestamp & " [" & level & "] " & msg

    ' Destination 1: Debug.Print (IDE output)
    If Config.LOG_TO_IMMEDIATE Then
        Debug.Print logEntry
    End If

    ' Destination 2: Excel Sheet (LOGS)
    If Config.LOG_TO_SHEET Then
        WriteToLogSheet logEntry, level
    End If
End Sub

' ============================================
' Write to LOGS sheet
' ============================================

Private Sub WriteToLogSheet(logEntry As String, level As String)
    On Error Resume Next

    Dim logsSheet As Worksheet
    Dim nextRow As Long

    ' Create LOGS sheet if doesn't exist
    Set logsSheet = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    If logsSheet Is Nothing Then
        Set logsSheet = ThisWorkbook.Worksheets.Add
        logsSheet.name = LOG_SHEET_NAME
        ' Header row
        logsSheet.Range("A1:D1").value = Array("Timestamp", "Level", "Message", "Details")
        logsSheet.Range("A1:D1").Font.Bold = True
        logsSheet.Tab.Color = 255
    End If

    ' Find next empty row
    nextRow = logsSheet.Cells(logsSheet.Rows.count, 1).End(xlUp).Row + 1

    ' Write log entry
    With logsSheet
        .Cells(nextRow, 1).value = Format(Now(), "YYYY-MM-DD HH:MM:SS")
        .Cells(nextRow, 2).value = level
        .Cells(nextRow, 3).value = logEntry

        ' Color code by level
        Select Case level
            Case "DEBUG"
                .Cells(nextRow, 2).Interior.Color = RGB(200, 200, 200)  ' Light gray
            Case "INFO"
                .Cells(nextRow, 2).Interior.Color = RGB(100, 200, 255)  ' Light blue
            Case "WARNING"
                .Cells(nextRow, 2).Interior.Color = RGB(255, 200, 100)  ' Orange
            Case "ERROR"
                .Cells(nextRow, 2).Interior.Color = RGB(255, 100, 100)  ' Light red
        End Select
    End With

    On Error GoTo 0
End Sub

' ============================================
' Initialize: Create LOGS sheet on startup
' ============================================

Public Sub InitializeLogging()
    On Error Resume Next
    Dim logsSheet As Worksheet
    Set logsSheet = ThisWorkbook.Worksheets(LOG_SHEET_NAME)

    If logsSheet Is Nothing Then
        Set logsSheet = ThisWorkbook.Worksheets.Add
        logsSheet.name = LOG_SHEET_NAME
        logsSheet.Range("A1:D1").value = Array("Timestamp", "Level", "Message", "Details")
        logsSheet.Range("A1:D1").Font.Bold = True
    End If

    LogInfo "Logging initialized"
    On Error GoTo 0
End Sub






