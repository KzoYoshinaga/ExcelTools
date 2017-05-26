Option Explicit

Public Sub import(lib)
	Dim fso, fin
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set fin = fso.OpentextFile(lib, 1)
	ExecuteGlobal fin.ReadAll
End Sub

import "ExelRunner"
import "AccessRunner"

Dim excelMacros
Set excelMacros = New ExcelRunnableCollection

Dim accessMacros
Set accessMacros = New AccessRunnableCollection

' 編集可 ここから ******************************************

Dim dbName
dbName = "マスタ.accdb"

With accessMacros
	.add "満車ログインポート"
	.add "営業所マスタインポート"
	.add "車名マスタインポート"
	.add "NEO予約ログインポート"
	.add "お断り合計月別"
	.add "お断り合計日別"
End With

With excelMacros
	.add "お断り合計.xlsm", "月別店舗別"
	.add "お断り合計.xlsm", "月別Ｒ店別"
	.add "お断り合計.xlsm", "日別表紙"
	.add "お断り合計.xlsm", "日別シート"
End With

' 編集可 ここまで ******************************************

Dim fileSystem
Set fileSystem = CreateObject("Scripting.FileSystemObject")

Dim currentPath
currentPath = fileSystem.getParentFolderName(WScript.ScriptFullName)

accessMacros.excecute currentPath, dbName
excelMacros.execute currentPath

WScript.Echo("All Complete")
