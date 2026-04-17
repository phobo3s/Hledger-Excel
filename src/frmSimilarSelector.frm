VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSimilarSelector 
   Caption         =   "UserForm1"
   ClientHeight    =   300
   ClientLeft      =   -252
   ClientTop       =   -984
   ClientWidth     =   144
   OleObjectBlob   =   "frmSimilarSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSimilarSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit

Private pFrmAnswer As String
Private pSelected As Variant

Property Get FrmAnswer()
    FrmAnswer = pFrmAnswer
End Property
Property Get GetSelected()
    GetSelected = pSelected
End Property

Private Sub btnCancel_Click()
    pFrmAnswer = "Cancel"
    Me.hide
End Sub

Private Sub btnUpdate_Click()
    pFrmAnswer = "Update"
    Me.hide
End Sub

Private Sub lbxSimilars_Change()
    Dim Selected() As Variant 'selected row values array
    ReDim Selected(lbxSimilars.columnCount - 1)
    Dim i As Long
    Dim j As Long
    For i = 0 To (lbxSimilars.ListCount - 1)
        If lbxSimilars.Selected(i) Then
            For j = LBound(Selected) To UBound(Selected)
                Selected(j) = lbxSimilars.List(i, j)
            Next j
            Exit For
        Else
        End If
    Next i
    ' fill textboxes
    Me.tbxDesc.Text = Selected(0)
    Me.tbxDate.Text = Me.lblDate.Caption
    Me.tbxToAcct.Text = Selected(1)
    Me.tbxAmount.Text = Me.lblAmount.Caption
    Me.cbxSpecial.Text = Selected(2)
    
End Sub

Private Sub UserForm_Activate()
    Me.height = 390
    Me.width = 420
    Me.left = Application.left + Application.width / 2
    Me.top = Application.top + Application.height / 2
    Me.cbxSpecial.List() = Array("", "Buy/Sell", "Dividend", "Interest")
End Sub

