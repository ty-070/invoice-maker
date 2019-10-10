Attribute VB_Name = "Module1"
Option Explicit

Sub 請求書作成()
    Dim table As Variant
    table = range("請求データ").ListObject.DataBodyRange
    
    Dim invoiceID As Variant
    invoiceID = filterDuplicates(column(table, 1))
    
    Dim invoiceList As collection
    Set invoiceList = New collection
    Dim i As Integer
    For i = 0 To UBound(invoiceID)
        Dim rowList As Variant
        rowList = findRow(table, invoiceID(i))
        
        Dim tempInv As Invoice
        Set tempInv = New Invoice
        tempInv.Setter (rowList)
        
        invoiceList.Add tempInv
    Next i
    
    Dim inv As Invoice
    For Each inv In invoiceList
        Dim doc As Word.Document
        Set doc = newInvoice()
        editWord doc, inv
    Next inv
End Sub

' tableから値がtargetと一致する要素をsearchColumnで指定された列の中から探し、該当する要素を含んでいる行を返します
Function findRow(ByVal table As Variant, target As Variant, Optional searchColumn As Integer = 1) As Variant
    Dim rowList As collection
    Set rowList = New collection
    
    Dim r As Integer
    For r = LBound(table, 1) To UBound(table, 1)
        If target = table(r, searchColumn) Then
            rowList.Add (row(table, r))
        End If
    Next r
    
    findRow = toArray(rowList)
End Function

' cを配列にして返します
' ******** 返り値の配列の最初のインデックスは0です *********
Function toArray(c As collection) As Variant
    ReDim arr(c.Count - 1) As Variant
    Dim i As Integer
    For i = 0 To c.Count - 1
        arr(i) = c.item(i + 1)
    Next i
    toArray = arr
End Function

' rから重複している要素を削除した配列を返します
Function filterDuplicates(ByVal arr As Variant) As Variant
    Dim unique As New collection
    Dim c As Variant
    On Error Resume Next
    For Each c In arr
        unique.Add c, CStr(c)
    Next c
    On Error GoTo 0
    filterDuplicates = toArray(unique)
End Function

' 請求書ファイルを完成させ、保存します
Private Sub editWord(ByRef doc As Word.Document, inv As Invoice)
    Dim table As Word.table
    With doc
        .Bookmarks("発行日").range.InsertAfter "発行日: " & Format(Date, "Long Date")
        .Bookmarks("宛名").range.InsertAfter inv.Name & " 御中"
        .Bookmarks("振込期限").range.InsertAfter Format(inv.Deadline, "Long Date")
        .Bookmarks("請求金額").range.InsertAfter "\ " & Format(inv.Sum * 1.1, "#,#")
        .Bookmarks("小計").range.InsertAfter "\ " & Format(inv.Sum, "#,#")
        .Bookmarks("消費税").range.InsertAfter "\ " & Format(inv.Sum * 0.1, "#,#")
        .Bookmarks("請求書番号").range.InsertAfter "No. " & inv.Id
        Set table = .Tables(3)
    End With
    
    Dim det As Details
    For Each det In inv.DetailsList
        With table.rows
            .Add
            With .last
                .Cells(1).range = det.Description
                .Cells(2).range = "\ " & Format(det.Price, "#,#")
                .Cells(3).range = Format(det.Quantity, "#,#")
                .Cells(4).range = "\ " & Format(det.Amount, "#,#")
            End With
        End With
    Next det
    
    table.rows(2).Delete
    
    doc.SaveAs2 (ThisWorkbook.Path & "\請求書_" & inv.Id & "_" & inv.Name)
End Sub

' 新しく請求書Wordファイルを作成します
Function newInvoice() As Word.Document
    Dim app As Word.Application
    Set app = CreateObject("Word.Application")
    app.Visible = True
    Dim templatePath As String
    templatePath = ThisWorkbook.Path & "\請求書テンプレート.dotx"
    Set newInvoice = app.Documents.Add(templatePath)
End Function

' 二次元配列arrのインデックスaの行を返します
Function row(ByVal table As Variant, a As Integer) As Variant
    Dim start As Integer
    start = LBound(table, 2)
    
    Dim last As Integer
    last = UBound(table, 2)
    
    ReDim c(start To last) As Variant
    
    Dim i As Integer
    For i = start To last
        c(i) = table(a, i)
    Next i
    
    row = c
End Function

' 二次元配列arrのインデックスbの列を返します
Function column(ByVal arr As Variant, b As Integer) As Variant
    Dim start As Integer
    start = LBound(arr, 1)
    
    Dim last As Integer
    last = UBound(arr, 1)
    
    ReDim c(start To last) As Variant
    
    Dim i As Integer
    For i = start To last
        c(i) = arr(i, b)
    Next i
    
    column = c
End Function
