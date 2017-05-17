VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tuple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' メンバ *****************************
Private values As New collection

' インスタンスコピー *****************
Public Property Get clone() As Tuple
    Dim r As New Tuple
    r.setValues = newCollection(values)
    Set clone = r
End Property

' プライベートコンストラクタ *********
Private Property Get newTuple(vs As collection) As Tuple
    Dim r As New Tuple
    r.setValues = newCollection(vs)
    Set newTuple = r
End Property

' イニシャライザ *********************
Public Property Let setValues(vs As collection)
    Set values = newCollection(vs)
End Property

' オブジェクト共通 ********************
Public Property Get toString()
    toString = "Tuple(" & join(values, ", ") & ")"
End Property

Private Property Get join(c As collection, delimiter As String)
    Dim r As String
    For i = 1 To c.count
        If i < c.count Then
            r = r & c(i).toString() & delimiter
        Else
            r = r & c(i).toString()
        End If
    Next
    join = r
End Property

Public Property Get equals(that As Tuple)
    If toString() = that.toString Then
        equals = True
    Else
        equals = False
    End If
End Property

' メソッド *****************************

' 値抜き出し
' @param s: Selection 抜き出す値の選択範囲
' @return Tuple 抜き出した値を持つ新しいタプル
Public Property Get trim(s As Selection) As Tuple
    Dim f As Long
    f = greater(1, s.getFrom())
    Dim l As Long
    l = less(s.getLast(), count())
    ' 抜き出し
    Dim vs As New collection
    For i = f To l
        vs.add values(i)
    Next
    Set trim = newTuple(vs)
End Property

' 2分割
' @param no: Long 分割する場所のインデックス
' @return TupleCollection 分割されたタプルを含むコレクション
Public Property Get split(no As Long) As Tuples2
    no = middle(1, no, EOB())
    Dim r As Tuples2
    Set r = newTuples2(trim(newSelection(0, no - 1)), _
                         trim(newSelection(no, EOB())))
    Set split = r
End Property

' 値挿入
' @param that: Tuple 挿入する値を持つタプル
' @param dest: Long 挿入先の要素番号
' @return Tuple 値が挿入された新しいタプル
Public Property Get insert(that As Tuple, dest As Long) As Tuple
    ' 挿入先修正
    Dim d As Long
    d = middle(1, dest, count() + 1)
    ' タプル分割
    Dim t2 As Tuples2
    Set t2 = split(d)
    ' タプル配置
    Dim t3 As New Tuples3
    t3.setLeft = t2.getLeft()
    t3.setMiddle = that.clone()
    t3.setRight = t2.getRight()
    ' タプルマージ
    Set insert = t3.marge()
End Property

' 値挿入リバース
Public Property Get insertR(that As Tuple, dest As Long) As Tuple
    Set insertR = that.insert(clone(), dest)
End Property

' 値削除
' @param s: Selection 削除する値の選択範囲
' @return Tuple 要素が削除された新しいタプル
Public Property Get remove(s As Selection) As Tuple
    Dim f As Long
    f = s.getFrom()
    Dim l As Long
    l = less(s.getLast(), count())
    ' 削除
    Dim vs As collection
    Set vs = newCollection(values)
    For i = f To l
        If i > 0 Then vs.remove f
    Next
    Set remove = newTuple(vs)
End Property

' 値移動
' @param s: Selection 移動する値の選択範囲
' @param dest: Long 移動先の番号
' @param Tuple 値が移動された新しいタプル
Public Property Get move(s As Selection, dest As Long) As Tuple
    Dim t As Tuple
    Set t = trim(s)
    Dim f As Long
    f = s.getFrom()
    Dim l As Long
    l = s.getLast()
    If dest = middle(f, dest, l) Then
        Set move = clone()
    ElseIf dest = less(f, dest) Then
        Dim tmp As Tuple
        Set tmp = remove(s)
        Set move = tmp.insert(t, dest)
    ElseIf dest = greater(l, dest) Then
        Dim tmp2 As Tuple
        Set tmp2 = insert(t, dest)
        Set move = tmp2.remove(s)
    Else
        Debug.Assert False
    End If
End Property

' 値貼り付け
' @param that: Tuple
' @param dest: Long
Public Property Get paste(that As Tuple, dest As Long) As Tuple
    Dim thisS As Selection
    Set thisS = newSelection(1, count())
    Dim thatS As Selection
    Set thatS = newSelection(1, that.count())
    Dim thisSelection As Selection
    Dim thatSelection As Selection
    Set thisSelection = thisS.andSelection(thatS, dest)
    Set thatSelection = thatS.andSelection(thisS, 2 - dest)
    Dim r As Tuple
    Set r = remove(thisSelection)
    Set r = r.insert(that.trim(thatSelection), dest)
    Set paste = r
End Property

' 値貼り付けリバース
Public Property Get pasteR(that As Tuple, dest As Long) As Tuple
    Set pasteR = that.paste(clone(), dest)
End Property

' マージ
' @param that: Tuple マージ対象タプル
' @return Tuple マージされた新しいタプル
Public Property Get marge(that As Tuple) As Tuple
    Dim vs As collection
    Set vs = newCollection(values)
    Dim i As Long
    For i = 1 To that.count()
        vs.add that.getValue(i)
    Next
    Set marge = newTuple(vs)
End Property

' マージリバース
Public Property Get margeR(that As Tuple) As Tuple
    Set margeR = that.marge(clone())
End Property

' 交互マージ
' 要素数の少ない方の要素数に切り詰められる
' @param that: Tuple 交互にマージするタプル
' @return Tuple 交互にマージされた新しいタプル
Public Property Get zip(that As Tuple) As Tuple
    c = lessCount(that)
    Dim vs As New collection
    Dim i As Long
    For i = 1 To c
        vs.add values(i)
        vs.add that.getValue(i)
    Next
    Set zip = newTuple(vs)
End Property

' 交互マージリバース
Public Property Get zipR(that As Tuple) As Tuple
    Set zipR = that.zip(clone())
End Property

' スキップ交互マージ
Public Property Get skipZip(that As Tuple, n As Long) As Tuple
    If n = 0 Then
        Set skipZip = clone()
        Exit Property
    End If
    c = less(count(), that.count() * n)
    Dim vs As New collection
    Dim i As Long
    For i = 1 To c
        vs.add values(i)
        If i Mod n = 0 Then
            vs.add that.getValue(i / n)
        End If
    Next
    Set skipZip = newTuple(vs)
End Property

' スキップ交互マージリバース
Public Property Get skipZipR(that As Tuple, n As Long) As Tuple
    Set skipZipR = that.skipZip(clone(), n)
End Property

' スキップインサート
Public Property Get skipInsert(v, n As Long) As Tuple
    Set skipInsert = skipZip(newTuple(newFilledCollection(v, count() / n + 1)), n)
End Property

' パディング左
Public Property Get paddingLeft(v, size As Long) As Tuple
   Set paddingLeft = newTuple(newFilledCollection(v, size - count)).marge(clone)
End Property

' パディング右
Public Property Get paddingRight(v, size As Long) As Tuple
    Set paddingRight = marge(newTuple(newFilledCollection(v, size - count)))
End Property

' リピート
Public Property Get repeat(times As Long) As Tuple
    Dim t As New Tuple
    Dim i As Long
    For i = 1 To times
        Set t = t.marge(clone())
    Next
    Set repeat = t
End Property


' 要素コレクションを取得
' @return Collection 要素のコレクション
Public Property Get getValues() As collection
    Set getValues = newCollection(values)
End Property

' 要素を取得
' @param no 取り出す要素番号
' @return 取り出す要素
Public Property Get getValue(no As Long)
    Set getValue = values(no)
End Property

' 要素数を取得
' @return Long 要素の数
Public Property Get count() As Long
    count = values.count()
End Property

' 最終要素間インデックスの取得
' @return Long 最終要素の右の要素間インデックス
Public Property Get EOB() As Long
    EOB = count() + 1
End Property

' プライベート ***********************

' 要素数の少ない方の要素数を取得
' @param that: Tuple 比較対象タプル
' @return Long 要素数が少ない方の要素数
Private Property Get lessCount(that As Tuple)
    If values.count < that.count Then
        lessCount = values.count
    Else
        lessCount = that.count
    End If
End Property

