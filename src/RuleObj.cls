VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RuleObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pactive As Boolean
Private pdescRuleType As String
Private pdescription As String
Private pamountOp As String
Private pamount As Double
Private paccount As String
Private ptoAccount As String
Private pnewDescription As String
Private pspecial As String
Private ppriority As Long

' --- active ---
Public Property Get Active() As Boolean
    Active = pactive
End Property

Public Property Let Active(ByVal value As Boolean)
    pactive = value
End Property

' --- descRuleType ---
Public Property Get DescRuleType() As String
    DescRuleType = pdescRuleType
End Property

Public Property Let DescRuleType(ByVal value As String)
    pdescRuleType = value
End Property

' --- description ---
Public Property Get Description() As String
    Description = pdescription
End Property

Public Property Let Description(ByVal value As String)
    pdescription = value
End Property

' --- amountOp ---
Public Property Get AmountOp() As String
    AmountOp = pamountOp
End Property

Public Property Let AmountOp(ByVal value As String)
    pamountOp = value
End Property

' --- amount ---
Public Property Get Amount() As Double
    Amount = pamount
End Property

Public Property Let Amount(ByVal value As Double)
    pamount = value
End Property

' --- account ---
Public Property Get account() As String
    account = paccount
End Property

Public Property Let account(ByVal value As String)
    paccount = value
End Property

' --- toAccount ---
Public Property Get ToAccount() As String
    ToAccount = ptoAccount
End Property

Public Property Let ToAccount(ByVal value As String)
    ptoAccount = value
End Property

' --- new description ---
Public Property Get NewDescription() As String
    NewDescription = pnewDescription
End Property

Public Property Let NewDescription(ByVal value As String)
    pnewDescription = value
End Property

' --- special ---
Public Property Get Special() As String
    Special = pspecial
End Property

Public Property Let Special(ByVal value As String)
    pspecial = value
End Property

' --- priority ---
Public Property Get Priority() As Long
    Priority = ppriority
End Property

Public Property Let Priority(ByVal value As Long)
    ppriority = value
End Property






