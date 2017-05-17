Attribute VB_Name = "Tuples2Test"
' Tuples2 **********************************************

' 新しいインスタンスが生成されること
Private Function createNewInstance()
    ' SetTestName
    testName = "createNewInstance()"

    ' Arrange
    Dim values As New collection
    values.add "a"
     
    ' Do
    Dim ts2 As New Tuples2
    ts2.setLeft = newTuple(values)
    ts2.setRight = newTuple(values)
       
    ' Verification
    result = equals(ts2.toString, "Tuples3(Tuple(a), Tuple(a))")
    verify result, testName
End Function

' 左タプルが取り出せること
Private Function getLeft()
    ' SetTestName
    testName = "getLeft()"

    ' Arrange
    Dim values1 As New collection
    values1.add "a"
    Dim values2 As New collection
    values2.add "b"
    
    Dim ts2 As New Tuples2
    ts2.setLeft = newTuple(values1)
    ts2.setRight = newTuple(values2)
   
    'Do
    Dim r As Tuple
    Set r = ts2.getLeft
    
    ' Verification
    result = equals(r.toString, "Tuple(a)")
    verify result, testName
End Function

' 右タプルが取り出せること
Private Function getRight()
    ' SetTestName
    testName = "getRight()"

    ' Arrange
    Dim values1 As New collection
    values1.add "a"
    Dim values2 As New collection
    values2.add "b"
    
    Dim ts2 As New Tuples2
    ts2.setLeft = newTuple(values1)
    ts2.setRight = newTuple(values2)
   
    ' Do
    Dim r As Tuple
    Set r = ts2.getRight()
    
    ' Verification
    result = equals(r.toString, "Tuple(b)")
    verify result, testName
End Function

' 結合されたタプルが取り出せること
Private Function marge()
    ' SetTestName
    testName = "marge()"

    ' Arrange
    Dim values1 As New collection
    values1.add "a"
    Dim values2 As New collection
    values2.add "b"
    
    Dim ts2 As New Tuples2
    ts2.setLeft = newTuple(values1)
    ts2.setRight = newTuple(values2)
    
    ' Do
    Dim r As Tuple
    Set r = ts2.marge()
    
    ' Verification
    result = equals(r.toString, "Tuple(a, b)")
    verify result, testName
End Function




