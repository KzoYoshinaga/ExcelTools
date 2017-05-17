Attribute VB_Name = "TableTest"
Private Function test()

    
    Dim t As New table
    
    ' �c�ɍ�����
    Sheets("pile").Cells.Clear
    Call t. _
          import(cell("A", "A1")). _
          pile(t.import(cell("B", "A1"))). _
          export(cell("pile", "A1")). _
          terminate
    
    ' 2��΂��ŉ��ɍ�����
    Sheets("zip").Cells.Clear
    Call t. _
          import(cell("A", "A1")). _
          skipZip(t.import(cell("B", "A1")), 2). _
          export(cell("zip", "A1")). _
          terminate
    
    ' �ύX�����R�ɍ����邱�Ƃ��o����
    Sheets("test").Cells.Clear
    Call t. _
          import(cell("A", "A1")). _
          insertRow(t.import(cell("B", "A1")), 2). _
          repeatRows(2). _
          repeatColumns(3). _
          export(cell("test", "A1")). _
          terminate


    ' ���ۂ̎g���� **********************************************
    Sheets("test2").Cells.Clear
    Dim shopClassTable As table
    Dim midasi As table
    Dim midasiZipped As table
    
    ' �d�Ȃ荇�����e�[�u�����쐬
    Set shopClassTable = t.import(cell("���v", "A1")). _
                            pile(t.import(cell("�O��", "A1"))). _
                            skipPile(t.import(cell("����", "A1")), 2)
    
    ' �����T�C�Y�̌��o������m��
    Set midasi = shopClassTable. _
                   trimByColumns(newSelection(1, 1)). _
                   repeatColumns(shopClassTable.columnCount)
    
    ' ���o����������
    Set midasiZipped = shopClassTable.zipR(midasi)
 
    ' 4��ڂ���o��
    Call midasiZipped. _
            trimByColumns(newSelection(4, midasiZipped.columnCount)). _
            export(cell("test2", "A1")). _
            terminate

End Function


