Attribute VB_Name = "BringWindowToFront"
Option Explicit
'https://www.youtube.com/watch?v=MV17DP40E4o
'Uzman Excel Youtube kanalı.

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Boolean
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
    Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Boolean
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
#End If

' Bir Sonraki Ekran Enum
Private Const GW_HWNDNEXT = 2
' Maksimum Ekran Enum
Const SW_SHOWMAXIMIZED = 3

Public Sub BringFront(ByVal partialCaption As String)
    Dim lhWndP As Long
    If GetHandleFromPartialCaption(lhWndP, partialCaption) = True Then
        If Not lhWndP = 0 Then
            BringWindowToTop (lhWndP)
            ShowWindow lhWndP, SW_SHOWMAXIMIZED
        End If
    End If
End Sub
Public Function GetHandleFromPartialCaption(ByRef lWnd As Long, ByVal sCaption As String) As Boolean
    Dim lhWndP As Long
    Dim sStr As String
    Dim ise As String
    GetHandleFromPartialCaption = False
    lhWndP = FindWindow(vbNullString, vbNullString)
    Do While Not lhWndP = 0
        sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
        GetWindowText lhWndP, sStr, Len(sStr)
        sStr = left$(sStr, Len(sStr) - 1)
        ise = IsWindowVisible(lhWndP)
        If Not InStr(1, sStr, sCaption, vbTextCompare) = 0 And ise = True Then
            GetHandleFromPartialCaption = True: lWnd = lhWndP: Exit Do
        End If
        lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
    Loop
End Function






