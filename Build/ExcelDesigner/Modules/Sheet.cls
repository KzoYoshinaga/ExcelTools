VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' メンバ **********************************
Private r As Range


' イニシャライザ *********************
Public Property Let setValues(ran As Range)
    Set r = ran
End Property

' オブジェクト共通 ********************
Public Property Get toString()
    toString = "SheetRange(" & r.Column & ", " & r.row & ")"
End Property

Public Property Get equals(that As Tuple)
    If toString() = that.toString Then
        equals = True
    Else
        equals = False
    End If
End Property

' 行数取得
' @return Long シートの行数
Public Property Get getRowCount() As Long
    If r.value = "" Then
        getRowCount = 0
    ElseIf r.Offset(1, 0).value = "" Then
        getRowCount = 1
    Else
        getRowCount = r.End(xlDown).row
    End If
End Property

' 列数取得
' @return Long シートの列数
Public Property Get getColumnCount() As Long
    If r.value = "" Then
        getColumnCount = 0
    ElseIf r.Offset(0, 1).value = "" Then
        getColumnCount = 1
    Else
        getColumnCount = r.End(xlToRight).Column
    End If
End Property


' 指定された行の値のコレクションを返す　=>　独自オブジェクトでラップする
' @param rowNo: Long 取得する値コレクションの行
' @return Collection 指定した行の値コレクション
Public Function getRowCollection(rowNo As Long) As collection
    Dim c As New collection
    For i = 1 To getColumnCount()
        c.add newElement(r.Offset(rowNo - 1, i - 1).value)
    Next
    Set getRowCollection = c
End Function

