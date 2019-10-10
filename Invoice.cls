VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 請求書クラス

' 請求書番号
Public Id As String

' 請求先の名前
Public Name As String

' 振込締切日
Public Deadline As Date

' Detailsのコレクション
Public DetailsList As collection

Private Sub Class_Initialize()
    Id = ""
    Name = ""
    Set DetailsList = New collection
End Sub

Public Function Setter(ByVal arr As Variant)
    Id = arr(0)(1)
    Name = arr(0)(2)
    Deadline = arr(0)(8)
    
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        Dim tmp As Details
        Set tmp = New Details
        tmp.Setter (arr(i))
        DetailsList.Add tmp
    Next i
End Function

' 合計金額(消費税を含まない)を返します
Public Function Sum() As Integer
    Dim tmp As Long
    tmp = 0
    Dim det As Details
    For Each det In DetailsList
        tmp = tmp + det.Amount
    Next det
    Sum = tmp
End Function
