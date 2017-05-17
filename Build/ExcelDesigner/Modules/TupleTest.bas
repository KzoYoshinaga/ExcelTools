Attribute VB_Name = "TupleTest"
' Tuple **********************************************

' �V�����C���X�^���X����������邱��
Function createNewInstance()
    ' SetTestName
    testName = "createNewInstance()"

    ' Arrange
    Dim values As New collection
    values.add newElement("a")
     
    ' Do
    Dim t As New Tuple
    t.setValues = values
    
    ' Verification
    result = equals(t.toString, "Tuple(a)")
    verify result, testName
End Function

' trim() ***************************************************

' (0, 0)�͈̗͂v�f�������o���Ȃ�����
Private Function trimNoValueWhenFrom0Last0()
    ' SetTestName
    testName = "trimNoValueWhenFrom0Last0()"

    ' Arrange
    Dim values As New collection
    values.add newElement("a")
    values.add newElement("b")
    
    Dim t As Tuple
    Set t = newTuple(values)
    
    Dim s As Selection
    Set s = newSelection(0, 0)
     
    ' Do
    Set t = t.trim(s)
    
    ' Verifivation
    result = equals(t.toString, "Tuple()")
    verify result, testName
End Function

' �v�f���ꌏ�����o���邱��
Private Function trim1Value()
    ' SetTestName
    testName = "trim1Value()"

    ' Arrange
    Dim values As New collection
    values.add newElement("a")
    values.add newElement("b")
     
    Dim t As Tuple
    Set t = newTuple(values)
    
    Dim s As Selection
    Set s = newSelection(1, 1)
     
    ' Do
    Set t = t.trim(s)
    
    ' Verifivation
    result = equals(t.toString, "Tuple(a)")
    verify result, testName
End Function

' �v�f�𕡐��������o���邱��
Private Function trimValues()
    ' SetTestName
    testName = "trimValues()"

    ' Arrange
    Dim values As New collection
    values.add newElement("a")
    values.add newElement("b")
    values.add newElement("c")
     
    Dim t As Tuple
    Set t = newTuple(values)
    
    Dim s As Selection
    Set s = newSelection(1, 2)
     
    ' Do
    Set t = t.trim(s)
    
    ' Verifivation
    result = equals(t.toString, "Tuple(a, b)")
    verify result, testName
End Function

' �^�v���O�ifrom, last �Ƃ���count ���傫���j�͈͂��w�肷��Ƌ�̃^�v�����A��
Private Function trimEmptyTupleWhenFromAndLastGreaterThanCount()
    ' SetTestName
    testName = "trimEmptyTupleWhenFromAndLastGreaterThanCount()"

    ' Arrange
    Dim values As New collection
    values.add newElement("a")
    values.add newElement("b")
    values.add newElement("c")
    values.add newElement("d")
     
    Dim t As Tuple
    Set t = newTuple(values)
    
    Dim s As Selection
    Set s = newSelection(5, 5)
     
    ' Do
    Set t = t.trim(s)
            
    ' Verifivation
    result = equals(t.toString, "Tuple()")
    verify result, testName
End Function

' split() ***************************************************

' �v�f��2�ɕ�������邱��
Private Function split2TuplesWhenInsideCount()
    ' SetTestName
    testName = "split2TuplesWhenNoInsideCount()"
    
    ' Arrange
    Dim values As New collection
    values.add newElement("a")
    values.add newElement("b")
    values.add newElement("c")
    values.add newElement("d")
     
    Dim t As New Tuple
    t.setValues = values
    
    ' Do
    Dim ts2 As Tuples2
    Set ts2 = t.split(4)
                       
    ' Verification
    result = equals(ts2.toString, "Tuples2(Tuple(a, b, c), Tuple(d))")
    verify result, testName
End Function

' �������v�f�ɂȂ邱��
Private Function rightTuplesEmptyWhenLargerCount()
    ' SetTestName
    testName = "rightTuplesEmptyWhenLargerCount()"
    
    ' Arrange
    Dim values As New collection
    values.add newElement("a")
    values.add newElement("b")
    values.add newElement("c")
    values.add newElement("d")
     
    Dim t As New Tuple
    t.setValues = values
    
    ' Do
    Dim ts2 As Tuples2
    Set ts2 = t.split(5)
            
    ' Verification
    result = equals(ts2.toString, "Tuples2(Tuple(a, b, c, d), Tuple())")
    verify result, testName
End Function

' �O������v�f�ɂȂ邱��
Private Function leftTuplesEmptyWhenLessCount()
    ' SetTestName
    testName = "leftTuplesEmptyWhenLessCount()"
    
    ' Arrange
    Dim values As New collection
    values.add newElement("a")
    values.add newElement("b")
    values.add newElement("c")
    values.add newElement("d")
     
    Dim t As New Tuple
    t.setValues = values
    
    ' Do
    Dim ts2 As Tuples2
    Set ts2 = t.split(-1)
        
    ' Verification
    result = equals(ts2.toString, "Tuples2(Tuple(), Tuple(a, b, c, d))")
    verify result, testName
End Function

' insert() ************************************************************

' �ړ����1��菬�����l���w�肵���ꍇ�A�擪�ɑ}������邱��
Private Function insertToTopWhenDestLessThan1()
    ' SetTestName
    testName = "insertToTopWhenDestLessThan1()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("1")
    values1.add newElement("2")
    values1.add newElement("3")
    values1.add newElement("4")
    values1.add newElement("5")
    
    Dim values2 As New collection
    values2.add newElement("a")
    values2.add newElement("b")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.insert(t2, -1)
        
    ' Verification
    result = equals(t.toString, "Tuple(a, b, 1, 2, 3, 4, 5)")
    verify result, testName
End Function

' �ړ����1���w�肵���ꍇ�A�擪�ɑ}������邱��
Function insertToTopWhenDestEquals1()
    ' SetTestName
    testName = "insertToTopWhenDestEquals1()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("1")
    values1.add newElement("2")
    values1.add newElement("3")
    values1.add newElement("4")
    values1.add newElement("5")
    
    Dim values2 As New collection
    values2.add newElement("a")
    values2.add newElement("b")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.insert(t2, 1)
    
    ' Verification
    result = equals(t.toString, "Tuple(a, b, 1, 2, 3, 4, 5)")
    verify result, testName
End Function

' �ړ����1���傫���A�^�v���T�C�Y��菬�����l���w�肵���ꍇ�A�w��ʒu�ɑ}������邱��
Private Function insertToDestWhenDestWasMiddleOfTuple()
    ' SetTestName
    testName = "insertToDestWhenDestWasMiddleOfTuple()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("1")
    values1.add newElement("2")
    values1.add newElement("3")
    values1.add newElement("4")
    values1.add newElement("5")
    
    Dim values2 As New collection
    values2.add newElement("a")
    values2.add newElement("b")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.insert(t2, 3)
    
    ' Verification
    result = equals(t.toString, "Tuple(1, 2, a, b, 3, 4, 5)")
    verify result, testName
End Function

' �ړ���Ƀ^�v���T�C�Y���w�肵���ꍇ�A�ŏI����1��������ɑ}������邱��
Function insertToPreLastWhenDestEqualsLast()
    ' SetTestName
    testName = "insertToPreLastWhenDestEqualsLast()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("1")
    values1.add newElement("2")
    values1.add newElement("3")
    values1.add newElement("4")
    values1.add newElement("5")
    
    Dim values2 As New collection
    values2.add newElement("a")
    values2.add newElement("b")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.insert(t2, 5)
    
    ' Verification
    result = equals(t.toString, "Tuple(1, 2, 3, 4, a, b, 5)")
    verify result, testName
End Function

' �ړ���Ƀ^�v���T�C�Y���傫�Ȓl���w�肵���ꍇ�A�ŏI��ɑ}���i�}�[�W�j����邱��
Function insertToLastWhenDestGreaterThanLast()
    ' SetTestName
    testName = "insertToLastWhenDestGreaterThanLast()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("1")
    values1.add newElement("2")
    values1.add newElement("3")
    values1.add newElement("4")
    values1.add newElement("5")
    
    Dim values2 As New collection
    values2.add newElement("a")
    values2.add newElement("b")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.insert(t2, 6)
    
    ' Verification
    result = equals(t.toString, "Tuple(1, 2, 3, 4, 5, a, b)")
    verify result, testName
End Function

' remove() *********************************************

' �v�f���ꌏ�폜�ł��邱��
Private Function removeValue()
    ' SetTestName
    testName = "removeValue()"

    ' Arrange
    Dim values As New collection
    values.add newElement("a")
    values.add newElement("b")
     
    Dim t As New Tuple
    t.setValues = values
    
    ' Do
    Set t = t.remove(newSelection(1, 1))
    
    ' Verifivation
    result = equals(t.toString, "Tuple(b)")
    verify result, testName
End Function

' �v�f���������폜�ł��邱��
Private Function removeValues()
    ' SetTestName
    testName = "removeValues()"

    ' Arrange
    Dim values As New collection
    values.add newElement("a")
    values.add newElement("b")
    values.add newElement("c")
     
    Dim t As New Tuple
    t.setValues = values
    
    ' Do
    Set t = t.remove(newSelection(1, 2))
    
    ' Verifivation
    result = equals(t.toString, "Tuple(c)")
    verify result, testName
End Function

' �v�f���ȏ�̌������w�肷��ƁA�J�n�ʒu�ȍ~�̑S�����폜�����
Private Function removeAllAfterFrom()
    ' SetTestName
    testName = "canRemoveAllAfterFrom()"

    ' Arrange
    Dim values As New collection
    values.add newElement("a")
    values.add newElement("b")
    values.add newElement("c")
    values.add newElement("d")
  
    Dim t As New Tuple
    t.setValues = values
    
    ' Do
    Set t = t.remove(newSelection(2, 100))
    
    ' Verifivation
    result = equals(t.toString, "Tuple(a)")
    verify result, testName
End Function


' move() ***********************************************

' �ړ��悪1��菬�����ꍇ�擪�Ɉړ������
Private Function moveToTopWhenDestLessThan1()
    'SetTestName
    testName = "moveToTopWhenDestLessThan1()"
    
    ' Arrange
    Dim values As New collection
    values.add newElement("1")
    values.add newElement("2")
    values.add newElement("3")
    values.add newElement("4")
    values.add newElement("5")
    values.add newElement("6")
    values.add newElement("7")
    values.add newElement("8")
    
    Dim t As Tuple
    Set t = newTuple(values)
    
    ' Do
    Set t = t.move(s:=newSelection(2, 3), dest:=-1)
        
    ' Verification
    result = equals(t.toString, "Tuple(2, 3, 1, 4, 5, 6, 7, 8)")
    verify result, testName
End Function

' �ړ��悪1�̏ꍇ�擪�Ɉړ������
Private Function moveToTopWhenDestEquals1()
    'SetTestName
    testName = "moveToTopWhenDestEquals1()"
    
    ' Arrange
    Dim values As New collection
    values.add newElement("1")
    values.add newElement("2")
    values.add newElement("3")
    values.add newElement("4")
    values.add newElement("5")
    values.add newElement("6")
    values.add newElement("7")
    values.add newElement("8")

    Dim t As Tuple
    Set t = newTuple(values)
    
    ' Do
    Set t = t.move(s:=newSelection(2, 3), dest:=1)
    
    ' Verification
    result = equals(t.toString, "Tuple(2, 3, 1, 4, 5, 6, 7, 8)")
    verify result, testName
End Function

' �ړ��悪�I��͈͓��Ȃ�Έړ����Ȃ�
Private Function moveSamePlaceWhenDestInsideSelection()
    'SetTestName
    testName = "moveSamePlaceWhenDestInsideSelection()"
    
    ' Arrange
    Dim values As New collection
    values.add newElement("1")
    values.add newElement("2")
    values.add newElement("3")
    values.add newElement("4")
    values.add newElement("5")
    values.add newElement("6")
    values.add newElement("7")
    values.add newElement("8")
  
    Dim t As Tuple
    Set t = newTuple(values)
    
    ' Do
    Set t = t.move(s:=newSelection(2, 5), dest:=3)
    
    ' Verification
    result = equals(t.toString, "Tuple(1, 2, 3, 4, 5, 6, 7, 8)")
    verify result, testName
End Function

' �ړ��悪(�S�̌���-�I������)���傫���ꍇ�ŏI�ԍ��Ɉړ������
Private Function moveToLastWhenDestGreaterThanAndEqualsRemovedTupleCount()
    'SetTestName
    testName = "moveToLastWhenDestGreaterThanAndEqualsRemovedTupleCount()"
    
    ' Arrange
    Dim values As New collection
    values.add newElement("1")
    values.add newElement("2")
    values.add newElement("3")
    values.add newElement("4")
    values.add newElement("5")
    values.add newElement("6")
    values.add newElement("7")
    values.add newElement("8")
  
    Dim t As Tuple
    Set t = newTuple(values)
    
    ' Do
    Set t = t.move(s:=newSelection(2, 3), dest:=9)
        
    ' Verification
    result = equals(t.toString, "Tuple(1, 4, 5, 6, 7, 8, 2, 3)")
    verify result, testName
End Function

' paste() *************************************************

' �O���\��t���ʒu���w�肵�A�͈͂��d�Ȃ�Ȃ��ꍇ�\��t���Ȃ�
Private Function pasteEmptyWhenSelectionForward()
    testName = "pasteEmptyWhenSelectionForward()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("1")
    values1.add newElement("2")
    values1.add newElement("3")
    values1.add newElement("4")
    values1.add newElement("5")
    values1.add newElement("6")
    values1.add newElement("7")
    values1.add newElement("8")
    
    Dim values2 As New collection
    values2.add newElement("a")
    values2.add newElement("b")
    values2.add newElement("c")
    values2.add newElement("d")
    values2.add newElement("e")
    
    Dim this As Tuple
    Set this = newTuple(values1)
    
    Dim that As Tuple
    Set that = newTuple(values2)
    
    Dim r As Tuple
    Set r = this.paste(that, -8)
                
    result = equals(r.toString, "Tuple(1, 2, 3, 4, 5, 6, 7, 8)")
    verify result, testName
End Function

' ����\��t���ʒu���w�肵�A�͈͂��d�Ȃ�Ȃ��ꍇ�\��t���Ȃ�
Private Function pasteEmptyWhenSelectionBackward()
    testName = "pasteEmptyWhenSelectionBackward()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("1")
    values1.add newElement("2")
    values1.add newElement("3")
    values1.add newElement("4")
    values1.add newElement("5")
    values1.add newElement("6")
    values1.add newElement("7")
    values1.add newElement("8")
    
    Dim values2 As New collection
    values2.add newElement("a")
    values2.add newElement("b")
    values2.add newElement("c")
    values2.add newElement("d")
    values2.add newElement("e")
    
    Dim this As Tuple
    Set this = newTuple(values1)
    
    Dim that As Tuple
    Set that = newTuple(values2)
    
    Dim r As Tuple
    Set r = this.paste(that, 10)
                
    result = equals(r.toString, "Tuple(1, 2, 3, 4, 5, 6, 7, 8)")
    verify result, testName
End Function

' �O���\��t���ʒu���w�肵�A�͈͂��d�Ȃ�Ȃ��ꍇ�\��t���Ȃ�(���E)
Private Function pasteEmptyWhenSelectionForwardLimit()
    testName = "pasteEmptyWhenSelectionForwardLimit()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("1")
    values1.add newElement("2")
    values1.add newElement("3")
    values1.add newElement("4")
    values1.add newElement("5")
    values1.add newElement("6")
    values1.add newElement("7")
    values1.add newElement("8")
    
    Dim values2 As New collection
    values2.add newElement("a")
    values2.add newElement("b")
    values2.add newElement("c")
    values2.add newElement("d")
    values2.add newElement("e")
    
    Dim this As Tuple
    Set this = newTuple(values1)
    
    Dim that As Tuple
    Set that = newTuple(values2)
    
    Dim r As Tuple
    Set r = this.paste(that, -4)
                
    result = equals(r.toString, "Tuple(1, 2, 3, 4, 5, 6, 7, 8)")
    verify result, testName
End Function

' ����\��t���ʒu���w�肵�A�͈͂��d�Ȃ�Ȃ��ꍇ�\��t���Ȃ�(���E)
Private Function pasteEmptyWhenSelectionBackwardLimit()
    testName = "pasteEmptyWhenSelectionBackwardLimit()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("1")
    values1.add newElement("2")
    values1.add newElement("3")
    values1.add newElement("4")
    values1.add newElement("5")
    values1.add newElement("6")
    values1.add newElement("7")
    values1.add newElement("8")
    
    Dim values2 As New collection
    values2.add newElement("a")
    values2.add newElement("b")
    values2.add newElement("c")
    values2.add newElement("d")
    values2.add newElement("e")
    
    Dim this As Tuple
    Set this = newTuple(values1)
    
    Dim that As Tuple
    Set that = newTuple(values2)
    
    Dim r As Tuple
    Set r = this.paste(that, 9)
                
    result = equals(r.toString, "Tuple(1, 2, 3, 4, 5, 6, 7, 8)")
    verify result, testName
End Function

' �O���\��t���ʒu���w�肵�A�͈͂��d�Ȃ镔�������\��t������邱��
Private Function pasteWhenSelectionForward()
    testName = "pasteWhenSelectionForward()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("1")
    values1.add newElement("2")
    values1.add newElement("3")
    values1.add newElement("4")
    values1.add newElement("5")
    values1.add newElement("6")
    values1.add newElement("7")
    values1.add newElement("8")
    
    Dim values2 As New collection
    values2.add newElement("a")
    values2.add newElement("b")
    values2.add newElement("c")
    values2.add newElement("d")
    values2.add newElement("e")
    
    Dim this As Tuple
    Set this = newTuple(values1)
    
    Dim that As Tuple
    Set that = newTuple(values2)
    
    Dim r As Tuple
    Set r = this.paste(that, -2)
                
    result = equals(r.toString, "Tuple(d, e, 3, 4, 5, 6, 7, 8)")
    verify result, testName
End Function

' ����\��t���ʒu���w�肵�A�͈͂��d�Ȃ镔�������\��t������邱��
Private Function pasteWhenSelectionBackward()
    testName = "pasteWhenSelectionBackward()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("1")
    values1.add newElement("2")
    values1.add newElement("3")
    values1.add newElement("4")
    values1.add newElement("5")
    values1.add newElement("6")
    values1.add newElement("7")
    values1.add newElement("8")
    
    Dim values2 As New collection
    values2.add newElement("a")
    values2.add newElement("b")
    values2.add newElement("c")
    values2.add newElement("d")
    values2.add newElement("e")
    
    Dim this As Tuple
    Set this = newTuple(values1)
    
    Dim that As Tuple
    Set that = newTuple(values2)
    
    Dim r As Tuple
    Set r = this.paste(that, 6)
                
    result = equals(r.toString, "Tuple(1, 2, 3, 4, 5, a, b, c)")
    verify result, testName
End Function

' �\��t���Ώۂ��̈���Ɏ��܂�Ƃ��A���ׂĂ̑Ώۂ��\��t������邱��
Private Function pasteWholeSelection()
    testName = "pasteWholeSelection()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("1")
    values1.add newElement("2")
    values1.add newElement("3")
    values1.add newElement("4")
    values1.add newElement("5")
    values1.add newElement("6")
    values1.add newElement("7")
    values1.add newElement("8")
    
    Dim values2 As New collection
    values2.add newElement("a")
    values2.add newElement("b")
    values2.add newElement("c")
    values2.add newElement("d")
    values2.add newElement("e")
    
    Dim this As Tuple
    Set this = newTuple(values1)
    
    Dim that As Tuple
    Set that = newTuple(values2)
    
    Dim r As Tuple
    Set r = this.paste(that, 2)
                
    result = equals(r.toString, "Tuple(1, a, b, c, d, e, 7, 8)")
    verify result, testName
End Function

' marge() ****************************************

' ���A�ΏۂƂ��ɋ�łȂ��Ƃ��A�Ώۂ��}�[�W����邱��
Private Function margeWhenOriginAndOppositAreNotEmpty()
    ' SetTestName
    testName = "margeWhenOriginAndOppositAreNotEmpty()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    
    Dim values2 As New collection
    values2.add newElement("d")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.marge(t2)
    
    ' Verification
    result = equals(t.toString, "Tuple(a, b, c, d)")
    verify result, testName
End Function

' ������̂Ƃ��A�Ώۂ��݂̂̃^�v�����Ԃ邱��
Private Function margeReturnOppositeWhenOriginWasEmpty()
    ' SetTestName
    testName = "margeReturnOppositeWhenOriginWasEmpty()"
    
    ' Arrange
    Dim values1 As New collection
    
    Dim values2 As New collection
    values2.add newElement("d")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.marge(t2)
    
    ' Verification
    result = equals(t.toString, "Tuple(d)")
    verify result, testName
End Function


' �Ώۂ���̂Ƃ��A���݂̂̃^�v�����Ԃ���邱��
Private Function margeReturnOriginWhenOppositeWasEmpty()
    ' SetTestName
    testName = "margeReturnOriginWhenOppositeWasEmpty()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    
    Dim values2 As New collection
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.marge(t2)
    
    ' Verification
    result = equals(t.toString, "Tuple(a, b, c)")
    verify result, testName
End Function

' zip() ***************************************************

' �����v�f���̃^�v�������݂Ƀ}�[�W����邱��
Private Function zipWith2SameSizeTupples()
    ' SetTestName
    testName = "zipWith2SameSizeTupples()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("c")
    values1.add newElement("e")
    
    Dim values2 As New collection
    values2.add newElement("b")
    values2.add newElement("d")
    values2.add newElement("f")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.zip(t2)
    
    ' Verification
    result = equals(t.toString, "Tuple(a, b, c, d, e, f)")
    verify result, testName
End Function

' ���^�v���̗v�f�����}�[�W�Ώۂ������Ȃ��ꍇ�A���v�f���Ő؂�l�߂��邱��
Private Function zipWithOriginSize()
   ' SetTestName
    testName = "zipWithOriginSize()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("c")
    
    Dim values2 As New collection
    values2.add newElement("b")
    values2.add newElement("d")
    values2.add newElement("f")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.zip(t2)
    
    ' Verification
    result = equals(t.toString, "Tuple(a, b, c, d)")
    verify result, testName
End Function

' �}�[�W�Ώۃ^�v���̗v�f�������^�v���������Ȃ��ꍇ�A�Ώۗv�f���Ő؂�l�߂��邱��
Private Function zipWithObjectiveSize()
    ' SetTestName
    testName = "zipWithOppositSize()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    
    Dim values2 As New collection
    values2.add newElement("b")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.zip(t2)
    
    ' Verification
    result = equals(t.toString, "Tuple(a, b)")
    verify result, testName
End Function

' skipZip ******************************************

Private Function skipZipReturnOriginWhen0()
    ' SetTestName
    testName = "skipZipReturnOriginWhen0()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    values1.add newElement("d")
    values1.add newElement("e")
    values1.add newElement("f")
    
    Dim values2 As New collection
    values2.add newElement("1")
    values2.add newElement("2")
    values2.add newElement("3")
    values2.add newElement("4")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.skipZip(t2, 0)
     
    ' Verification
    result = equals(t.toString, "Tuple(a, b, c, d, e, f)")
    verify result, testName
End Function

Private Function skipZipWasSameResaltWhen1()
    ' SetTestName
    testName = "skipZipWasSameResaltWhen1()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    values1.add newElement("d")
    values1.add newElement("e")
    values1.add newElement("f")
    
    Dim values2 As New collection
    values2.add newElement("1")
    values2.add newElement("2")
    values2.add newElement("3")
    values2.add newElement("4")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.skipZip(t2, 1)
          
    ' Verification
    result = equals(t.toString, "Tuple(a, 1, b, 2, c, 3, d, 4)")
    verify result, testName
End Function

Private Function skipZipWhenGreaterThan2()
    ' SetTestName
    testName = "skipZipWhenGreaterThan2()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    values1.add newElement("d")
    values1.add newElement("e")
    values1.add newElement("f")
    
    Dim values2 As New collection
    values2.add newElement("1")
    values2.add newElement("2")
    values2.add newElement("3")
    values2.add newElement("4")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.skipZip(t2, 2)
        
    ' Verification
    result = equals(t.toString, "Tuple(a, b, 1, c, d, 2, e, f, 3)")
    verify result, testName
End Function

Private Function skipZipReturnOriginWhenOverOriginCount()
     ' SetTestName
    testName = "skipZipReturnOriginWhenOverOriginCount()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    values1.add newElement("d")
    values1.add newElement("e")
    values1.add newElement("f")
    
    Dim values2 As New collection
    values2.add newElement("1")
    values2.add newElement("2")
    values2.add newElement("3")
    values2.add newElement("4")
     
    Dim t1 As New Tuple
    t1.setValues = values1
    
    Dim t2 As New Tuple
    t2.setValues = values2
    
    ' Do
    Dim t As Tuple
    Set t = t1.skipZip(t2, 7)
        
    ' Verification
    result = equals(t.toString, "Tuple(a, b, c, d, e, f)")
    verify result, testName
End Function

' skipInsert() ******************************************

Private Function skipInsertTest()
     ' SetTestName
    testName = "skipZipReturnOriginWhenOverOriginCount()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    values1.add newElement("d")
    values1.add newElement("e")
    values1.add newElement("f")
    values1.add newElement("g")
    
    Dim t1 As New Tuple
    t1.setValues = values1
    
    ' Do
    Dim t As Tuple
    Set t = t1.skipInsert(newElement("9"), 4)
    
    Debug.Print t.toString
        
    ' Verification
    result = equals(t.toString, "Tuple(a, b, c, d, 9, e, f, g)")
    verify result, testName
End Function

' paddingLeft() *********************************************

Private Function paddingLeftReturnOriginWhenLessThanAndEqualsCount()
     ' SetTestName
    testName = "paddingLeftReturnOriginWhenLessThanAndEqualsCount()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    values1.add newElement("d")
    
    Dim t1 As New Tuple
    t1.setValues = values1
    
    ' Do
    Dim t As Tuple
    Set t = t1.paddingLeft(newElement("0"), 4)
        
    ' Verification
    result = equals(t.toString, "Tuple(a, b, c, d)")
    verify result, testName
End Function

Private Function paddingLeftGreaterThanCount()
     ' SetTestName
    testName = "paddingLeftReturnOriginWhenLessThanAndEqualsCount()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    values1.add newElement("d")
    
    Dim t1 As New Tuple
    t1.setValues = values1
    
    ' Do
    Dim t As Tuple
    Set t = t1.paddingLeft(newElement("0"), 9)
    
    ' Verification
    result = equals(t.toString, "Tuple(0, 0, 0, 0, 0, a, b, c, d)")
    verify result, testName
End Function

' paddingRight() **********************************************

Private Function paddingRightReturnOriginWhenLessThanAndEqualsCount()
     ' SetTestName
    testName = "paddingRightReturnOriginWhenLessThanAndEqualsCount()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    values1.add newElement("d")
    
    Dim t1 As New Tuple
    t1.setValues = values1
    
    ' Do
    Dim t As Tuple
    Set t = t1.paddingRight(newElement("0"), 4)
        
    ' Verification
    result = equals(t.toString, "Tuple(a, b, c, d)")
    verify result, testName
End Function

Private Function paddingRightGreaterThanCount()
    ' SetTestName
    testName = "paddingRightReturnOriginWhenLessThanAndEqualsCount()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    values1.add newElement("d")
    
    Dim t1 As New Tuple
    t1.setValues = values1
    
    ' Do
    Dim t As Tuple
    Set t = t1.paddingRight(newElement("0"), 9)
    
    ' Verification
    result = equals(t.toString, "Tuple(a, b, c, d, 0, 0, 0, 0, 0)")
    verify result, testName
End Function


' repeat() ********************************************************

Public Function repeatReturnOriginWhenLessThanAndEqual1()
    ' SetTestName
    testName = "repeatReturnOriginWhenLessThanAndEqual1()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    values1.add newElement("d")
    
    Dim t1 As New Tuple
    t1.setValues = values1
    
    ' Do
    Dim t As Tuple
    Set t = t1.repeat(1)
    
    ' Verification
    result = equals(t.toString, "Tuple(a, b, c, d)")
    verify result, testName
End Function

Public Function repeatWhenGreaterThan1()
    ' SetTestName
    testName = "repeatWhenGreaterThan1()"
    
    ' Arrange
    Dim values1 As New collection
    values1.add newElement("a")
    values1.add newElement("b")
    values1.add newElement("c")
    values1.add newElement("d")
    
    Dim t1 As New Tuple
    t1.setValues = values1
    
    ' Do
    Dim t As Tuple
    Set t = t1.repeat(2)
    
    ' Verification
    result = equals(t.toString, "Tuple(a, b, c, d, a, b, c, d)")
    verify result, testName
End Function


Function TupleTest()
    Debug.Print ""
    Debug.Print "test start"
    
    createNewInstance
    
    ' trim()
    trimNoValueWhenFrom0Last0
    trim1Value
    trimValues
    trimEmptyTupleWhenFromAndLastGreaterThanCount
    
    ' split()
    split2TuplesWhenInsideCount
    rightTuplesEmptyWhenLargerCount
    leftTuplesEmptyWhenLessCount
    
    ' insert()
    insertToTopWhenDestLessThan1
    insertToTopWhenDestEquals1
    insertToDestWhenDestWasMiddleOfTuple
    insertToPreLastWhenDestEqualsLast
    insertToLastWhenDestGreaterThanLast

    ' remove()
    removeValue
    removeValues
    removeAllAfterFrom
        
    ' move()
    moveToTopWhenDestLessThan1
    moveToTopWhenDestEquals1
    moveSamePlaceWhenDestInsideSelection
    moveToLastWhenDestGreaterThanAndEqualsRemovedTupleCount
    
    ' paste()
    pasteEmptyWhenSelectionForward
    pasteEmptyWhenSelectionBackwardLimit
    pasteEmptyWhenSelectionForward
    pasteEmptyWhenSelectionForwardLimit
    pasteWhenSelectionBackward
    pasteWhenSelectionForward
    pasteWholeSelection
    
    ' marge()
    margeWhenOriginAndOppositAreNotEmpty
    margeReturnOppositeWhenOriginWasEmpty
    margeReturnOriginWhenOppositeWasEmpty
    
     ' zip()
    zipWith2SameSizeTupples
    zipWithOriginSize
    zipWithObjectiveSize
    
    ' skipZip()
    skipZipReturnOriginWhen0
    skipZipWasSameResaltWhen1
    skipZipWhenGreaterThan2
    skipZipReturnOriginWhenOverOriginCount
    
    ' paddingLeft()
    paddingLeftReturnOriginWhenLessThanAndEqualsCount
    paddingLeftGreaterThanCount
    
    ' paddingRight()
    paddingRightReturnOriginWhenLessThanAndEqualsCount
    paddingRightGreaterThanCount
    
    ' repeat()
    repeatReturnOriginWhenLessThanAndEqual1
    repeatWhenGreaterThan1
    
End Function
