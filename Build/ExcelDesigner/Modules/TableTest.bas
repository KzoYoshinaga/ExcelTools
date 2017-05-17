Attribute VB_Name = "TableTest"
Private Function test()

    
    Dim t As New table
    
    ' 縦に混ぜる
    Sheets("pile").Cells.Clear
    Call t. _
          import(cell("A", "A1")). _
          pile(t.import(cell("B", "A1"))). _
          export(cell("pile", "A1")). _
          terminate
    
    ' 2つ飛ばしで横に混ぜる
    Sheets("zip").Cells.Clear
    Call t. _
          import(cell("A", "A1")). _
          skipZip(t.import(cell("B", "A1")), 2). _
          export(cell("zip", "A1")). _
          terminate
    
    ' 変更を自由に混ぜることが出来る
    Sheets("test").Cells.Clear
    Call t. _
          import(cell("A", "A1")). _
          insertRow(t.import(cell("B", "A1")), 2). _
          repeatRows(2). _
          repeatColumns(3). _
          export(cell("test", "A1")). _
          terminate


    ' 実際の使い方 **********************************************
    Sheets("test2").Cells.Clear
    Dim shopClassTable As table
    Dim midasi As table
    Dim midasiZipped As table
    
    ' 重なり合ったテーブルを作成
    Set shopClassTable = t.import(cell("合計", "A1")). _
                            pile(t.import(cell("前日", "A1"))). _
                            skipPile(t.import(cell("当日", "A1")), 2)
    
    ' 同じサイズの見出し列を確保
    Set midasi = shopClassTable. _
                   trimByColumns(newSelection(1, 1)). _
                   repeatColumns(shopClassTable.columnCount)
    
    ' 見出しを混ぜる
    Set midasiZipped = shopClassTable.zipR(midasi)
 
    ' 4列目から出力
    Call midasiZipped. _
            trimByColumns(newSelection(4, midasiZipped.columnCount)). _
            export(cell("test2", "A1")). _
            terminate

End Function


