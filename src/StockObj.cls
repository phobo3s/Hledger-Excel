VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StockObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'stock Code
Private pKod As String 'as Stock ShortCode
'positions dictionaries
Private pPositions As Dictionary 'positions dictionary
Private pClosedPositions As Dictionary 'closed positions dictionary
' one position object
Private pPosition As PositionObj 'position Object
Private pCumulativeStockCount As Double 'cumulativeStock Count
Private pPositionNumber As Integer
Private pClosedPositionNumber As Integer
Private pOpenPositionCount As Integer

Private Sub Class_Initialize()
    If pPositions Is Nothing Then Set pPositions = New Dictionary
    If pClosedPositions Is Nothing Then Set pClosedPositions = New Dictionary
End Sub

Private Sub Class_Terminate()
    'If Not pPositions Is Nothing Then Set pPositions = Nothing
    'If Not pClosedPositions Is Nothing Then Set pClosedPositions = Nothing
End Sub

' Let get functions
Public Property Let code(value As String)
    pKod = value
End Property
Public Property Get code() As String
    code = pKod
End Property
Public Function CumulativeStockCount() As Double
    Dim aPos As Variant
    For Each aPos In pPositions.Items
        CumulativeStockCount = CumulativeStockCount + aPos.OpenQuantity
    Next aPos
End Function
Public Property Get OpenCount() As String
    OpenCount = pPositions.Count
End Property
Public Property Get ClosedCount() As String
    ClosedCount = pClosedPositions.Count
End Property

'create a position
Public Function OpenPosition(ddate As Date, Price As Double, quant As Double)
    quant = Math.Abs(quant)
    Price = Math.Abs(Price)
    Dim Position As New PositionObj
    pPositionNumber = pPositionNumber + 1
    Position.PosNum = pPositionNumber
    Position.OpenDate = ddate
    Position.OpenPrice = Price
    Position.OpenQuantity = quant
    pPositions.Add Position.PosNum, Position
End Function
'call open position
Public Property Get Position(PossNum As Integer) As PositionObj
    Set Position = pPositions.Items(PossNum - 1) '-1 because i want to positions to be start from 1
End Property
'call closed position
Public Property Get ClosedPosition(PossNum As Integer) As PositionObj
    Set ClosedPosition = pClosedPositions.Items(PossNum - 1) '-1 because i want to positions to be start from 1
End Property
'close position with FIFO
Public Function ClosePosition(ddate As Date, Price As Double, quant As Double) As PositionObj
    'close position according to FIFO principle.
    Dim Position As PositionObj
    Set Position = pPositions.Items(0) 'first position.
    Price = Math.Abs(Price)
    quant = Math.Abs(quant)
    Dim tempQuant As Double
    If quant > Position.OpenQuantity Then
        tempQuant = Position.OpenQuantity
        Position.PositionState = "Closed"
        Position.CloseDate = ddate
        Position.ClosePrice = tempQuant * (Price / quant)
        Position.CloseQuantity = tempQuant
        pPositions.Remove Position.PosNum
        pClosedPositionNumber = pClosedPositionNumber + 1
        pClosedPositions.Add pClosedPositionNumber, Position
        ClosePosition ddate, Price - (tempQuant * (Price / quant)), quant - tempQuant
    ElseIf quant < Position.OpenQuantity Then
        'close part of position
        tempQuant = Position.OpenQuantity
        Dim closingPosition As New PositionObj
        closingPosition.PosNum = Position.PosNum
        closingPosition.CloseDate = ddate
        closingPosition.ClosePrice = Price
        closingPosition.CloseQuantity = quant
        closingPosition.OpenDate = Position.OpenDate
        closingPosition.OpenPrice = Position.OpenPrice
        closingPosition.OpenPrice = quant * (Position.OpenPrice / Position.OpenQuantity)
        closingPosition.OpenQuantity = quant
        closingPosition.PositionState = "Closed"
        pClosedPositionNumber = pClosedPositionNumber + 1
        pClosedPositions.Add pClosedPositionNumber, closingPosition
        'change the remanining open position
        Position.OpenPrice = (tempQuant - quant) * (Position.OpenPrice / Position.OpenQuantity)
        Position.OpenQuantity = tempQuant - quant
    ElseIf quant = Position.OpenQuantity Then
        Position.PositionState = "Closed"
        Position.CloseDate = ddate
        Position.ClosePrice = Price
        Position.CloseQuantity = quant
        pPositions.Remove Position.PosNum
        pClosedPositionNumber = pClosedPositionNumber + 1
        pClosedPositions.Add pClosedPositionNumber, Position
    Else
        'error?!?
    End If
End Function






