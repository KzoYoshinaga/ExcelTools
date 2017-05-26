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
	.add "満車ログ月別"
	.add "満車ログ日別"
End With

With excelMacros
	.add "満車ログ.xlsx", "日別表紙"
	.add "満車ログ.xlsx", "日別シート"
	.add "満車ログ.xlsx", "月別店舗別"
	.add "満車ログ.xlsx", "月別Ｒ店別"
End With

' 編集可 ここまで ******************************************

Dim fileSystem
Set fileSystem = CreateObject("Scripting.FileSystemObject")

Dim currentPath
currentPath = fileSystem.getParentFolderName(WScript.ScriptFullName)

accessMacros.excecute currentPath, dbName
excelMacros.execute currentPath

WScript.Echo("All Complete")
