VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' メンバ *****************************
Private f As Long
Private l As Long

' コンストラクタ *********************
Public Sub class_initialize()
    f = -2147483647
    l = 2147483647
End Sub

' イニシャライザ *********************
Public Property Let setFrom(from As Long)
    f = less(from, l)
    l = greater(f, from)
End Property

Public Property Let setLast(last As Long)
    l = greater(last, f)
    f = less(f, last)
End Property

' オブジェクト共通 ********************
Public Property Get toString()
    toString = "Selection(" & f & ", " & l & ")"
End Property

Public Property Get equals(that As Tuple)
    If toString() = that.toString Then
        equals = True
    Else
        equals = False
    End If
End Property

' メソッド ***************************

Public Property Get getFrom()
    getFrom = f
End Property

Public Property Get getLast()
    getLast = l
End Property

' 積範囲
Public Property Get andSelection(that As Selection, dest As Long) As Selection
    Dim from As Long
    from = middle(dest, getFrom(), that.getLast() + dest - 1)
    Dim last As Long
    last = middle(dest, dest + that.getLast() - 1, getLast())
    Set andSelection = newSelection(from, last)
End Property


