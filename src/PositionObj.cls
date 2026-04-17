VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PositionObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPosNum As Integer 'position number
Private popenDate As Date
Private popenPrice As Double
Private popenQuantity As Double
Private pcloseDate As Date
Private pclosePrice As Double
Private pcloseQuantity As Double
Private ppositionState As String

Private Sub Class_Initialize()
    ppositionState = "Open"
End Sub

' Let get functions
'OpenDate
Public Property Let PosNum(value As Integer)
    pPosNum = value
End Property
Public Property Get PosNum() As Integer
    PosNum = pPosNum
End Property
'OpenDate
Public Property Let OpenDate(value As Date)
    popenDate = value
End Property
Public Property Get OpenDate() As Date
    OpenDate = popenDate
End Property
'OpenPrice
Public Property Let OpenPrice(value As Double)
    popenPrice = value
End Property
Public Property Get OpenPrice() As Double
    OpenPrice = popenPrice
End Property
'OpenQuantity
Public Property Let OpenQuantity(value As Double)
    popenQuantity = value
End Property
Public Property Get OpenQuantity() As Double
    OpenQuantity = popenQuantity
End Property
'CloseDate
Public Property Let CloseDate(value As Date)
    pcloseDate = value
End Property
Public Property Get CloseDate() As Date
    CloseDate = pcloseDate
End Property
'ClosePrice
Public Property Let ClosePrice(value As Double)
    pclosePrice = value
End Property
Public Property Get ClosePrice() As Double
    ClosePrice = pclosePrice
End Property
'CloseQuantity
Public Property Let CloseQuantity(value As Double)
    pcloseQuantity = value
End Property
Public Property Get CloseQuantity() As Double
    CloseQuantity = pcloseQuantity
End Property
Public Property Let PositionState(value As String)
    ppositionState = value
End Property
Public Property Get PositionState() As String
    PositionState = ppositionState
End Property






