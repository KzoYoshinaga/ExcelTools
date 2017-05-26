Option Explicit

Public Sub import(lib)
	Dim fso, fin
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set fin = fso.OpenTextFile(lib, 1)
	ExecuteGlobal fin.ReadAll
End Sub

import "ExcelRunner"
import "AccessRunner"

Dim excelMacros
Set excelMacros = New ExcelRunnableCollection

Dim accessMacros
Set accessMacros = New AccessRunnableCollection

' 編集可 ここから ******************************************

Dim dbName
dbName = "マスタ.accdb"

accessMacros.add "満車ログインポート"
accessMacros.add "営業所マスタインポート"
accessMacros.add "車名マスタインポート"
accessMacros.add "NEO予約ログインポート"
accessMacros.add "お断り合計月別"
accessMacros.add "お断り合計日別"

excelMacros.add "お断り合計.xlsm", "月別店舗別"
excelMacros.add "お断り合計.xlsm", "月別Ｒ店別"
excelMacros.add "お断り合計.xlsm", "日別表紙"
excelMacros.add "お断り合計.xlsm", "日別シート"

' 編集可 ここまで ******************************************

Dim fileSystem
Set fileSystem = CreateObject("Scripting.FileSystemObject")

Dim currentPath
currentPath = fileSystem.getParentFolderName(WScript.ScriptFullName)

accessMacros.execute currentPath, dbName
excelMacros.execute currentPath

WScript.Echo("All Complete")
