VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDescription 
   Caption         =   "Description Selection"
   ClientHeight    =   1536
   ClientLeft      =   24
   ClientTop       =   120
   ClientWidth     =   1872
   OleObjectBlob   =   "frmDescription.frx":0000
End
Attribute VB_Name = "frmDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit

Private pFrmAnswer As String

Property Get FrmAnswer()
    FrmAnswer = pFrmAnswer
End Property


'BUTTONS
'--------
Private Sub btnCancel_Click()
    pFrmAnswer = "Cancel"
    Me.hide
End Sub
Private Sub btnNo_Click()
    pFrmAnswer = "No"
    Me.hide
End Sub
Private Sub btnYes_Click()
    If tbxEditDesc = "" Then
        pFrmAnswer = "Yes"
    Else
        pFrmAnswer = tbxEditDesc.Text
    End If
    Me.hide
End Sub

'
'--------
Private Sub UserForm_Activate()
    Me.height = 230
    Me.width = 275
    Me.left = Application.left + Application.width / 2
    Me.top = Application.top + Application.height / 2
End Sub

