Attribute VB_Name = "Factory"

Function newTuple(c As collection) As Tuple
    Dim r As New Tuple
    r.setValues = c
    Set newTuple = r
End Function

Function newSelection(from As Long, last As Long) As Selection
    Dim r As New Selection
    r.setFrom = from
    r.setLast = last
    Set newSelection = r
End Function

Function newTuples2(left As Tuple, right As Tuple) As Tuples2
    Dim r As New Tuples2
    r.setLeft = left
    r.setRight = right
    Set newTuples2 = r
End Function

Function newTuples3(left As Tuple, middle As Tuple, right As Tuple) As Tuples3
    Dim r As New Tuples3
    r.setLeft = left
    r.setMiddle = middle
    r.setRight = right
    Set newTuples3 = r
End Function

Function newSheet(ran As Range) As sheet
    Dim r As New sheet
    r.setValues = ran
    Set newSheet = r
End Function

Function newElement(v) As Element
    Dim r As New Element
    r.setValue = v
    Set newElement = r
End Function

Function newTable(tuples As Tuple, h As Tuple) As table
    Dim r As New table
    r.setTuples = tuples
    r.setHeader = h
    Set newTable = r
End Function

Function newCollection(c As collection) As collection
    Dim r As New collection
    For Each v In c
        r.add v
    Next
    Set newCollection = r
End Function

Function newFilledCollection(v, size As Long) As collection
    Dim r As New collection
    Dim i As Long
    For i = 1 To size
        r.add v
    Next
    Set newFilledCollection = r
End Function
