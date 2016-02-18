VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Function RemoveRows()
    Dim Number As Integer
    Dim row As Integer
    row = 0
    Number = 625
    Do While row <= Number
        If (ActiveCell.Borders(xlEdgeLeft).LineStyle <> 1) Then
            ActiveSheet.Rows(ActiveCell.row).Delete
            row = row + 1
        Else
            range("A" & CStr(ActiveCell.row + ActiveCell.MergeArea.Rows.Count)).Activate
        End If
    Loop
End Function

Public Function ParseMuhProperties(sth As range) As Collection
    Dim ObjectProperties As New Collection, cols As range
    For Each cols In sth.Columns
        If cols.Value <> "" Then
            ObjectProperties.Add (cols.Value)
        End If
    Next cols
    Set ParseMuhProperties = ObjectProperties
End Function

Public Function ParseMuhRange(sth As range) As Collection
    Dim Rows As range, CollectionofCollections As New Collection
    For Each Rows In sth.Rows
        CollectionofCollections.Add ParseMuhProperties(Rows)
    Next Rows
    Set ParseMuhRange = CollectionofCollections
End Function

Public Function ParseJson(Properties As Collection, Values As Collection)
    Dim row As Collection, rowindex, column, json As String
    rowindex = 0
    For Each row In Values
        json = json & Chr(123)
        Dim colcount As Integer
        colcount = 0
        For Each column In row
            json = json & Chr(34) & Properties.Item(colcount + 1) & Chr(34)
            json = json & Chr(58) & Chr(34) & column & Chr(34)
            colcount = colcount + 1
            If colcount <> row.Count Then
                json = json & Chr(44) & vbNewLine
            End If
        Next column
        json = json & Chr(125)
        rowindex = rowindex + 1
        If Values.Count <> rowindex Then
            json = json & Chr(44) & vbNewLine
        End If
    Next row
    If Values.Count <> 1 Then
        json = Chr(91) & json & Chr(93)
    End If
    ParseJson = json
End Function

Public Function ParseMuhJson(text As String)
    Dim destination As Object
    Set destination = CreateObject("ADODB.Stream")
    destination.Type = 2
    destination.Charset = "utf-8"
    destination.Open
    destination.writeText text
    destination.savetofile Application.ActiveWorkbook.FullName & ".json", 2
    destination.Close
End Function

Sub stuff(Optional name As String = "")
    Dim Properties As Collection, Values As Collection
    Set Properties = ParseMuhProperties(Application.InputBox("Select Property Range", "Property Range", Type:=8))
    Set Values = ParseMuhRange(Application.InputBox("Select Property Range", "Property Range", Type:=8))
    name = name & vbNewLine & ParseJson(Properties, Values)
    If MsgBox("Add more Data to parse?", vbYesNo, "Add more data") = vbYes Then
        stuff (name & ",")
    Else
        ParseMuhJson (name)
    End If
End Sub

Sub start()
    Dim something As Integer
    stuff ("")
End Sub

