VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 請求書明細のクラス

' 品名
Public Description As String

' 単価
Public Price As Long

' 数量
Public Quantity As Long

' 金額を返します
Public Function Amount() As Long
    Amount = Price * Quantity
End Function

Public Function Setter(ByVal arr As Variant)
    Description = arr(3)
    Price = arr(4)
    Quantity = arr(5)
End Function

