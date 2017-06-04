Dim root 
if WScript.Arguments.Count <> 1 then
	Dim fileSystem
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	root = fileSystem.getParentFolderName(WScript.ScriptFullName)
else
	root = WScript.Arguments(0)
end if

Dim excelMacros
Set excelMacros = New ExcelRunnableCollection

Dim accessMacros
Set accessMacros = New AccessRunnableCollection

' �ҏW�� �������� ******************************************

db = "master.accdb"

accessMacros.add "���ԃ��O�C���|�[�g"
accessMacros.add "�c�Ə��}�X�^�C���|�[�g"
accessMacros.add "�Ԗ��}�X�^�C���|�[�g"
accessMacros.add "NEO�\�񃍃O�C���|�[�g"
accessMacros.add "���f�荇�v����"
accessMacros.add "���f�荇�v����"


book = "view.xlsm"

excelMacros.add "���ʓX�ܕ�"
excelMacros.add "����R�X��"
excelMacros.add "���ʕ\��"
excelMacros.add "���ʃV�[�g"

' �ҏW�� �����܂� ******************************************

accessMacros.execute root & "\Tools", db
excelMacros.execute root & "\Tools", book

WScript.Echo("All Complete")



' Class ****************************************************

Class ExcelRunnable
	Private macroName
	Private no

	Public Property Get setProperties(m, n)
		macroName = m
		no = n
	End Property

	Public Property Get run(excel, total)

		WScript.Echo( "(" & no & "/" & total & ") Book=" & bookName & " Macro=" & macroName & " Start" )

		result = excel.Application.Run(macroName)
		If result = 0 Then
			WScript.Echo( "(" & no & "/" & total & ") Book=" & bookName & " Macro=" & macroName & " Complete" )
			run = True
		Else
			WScript.Echo( "(" & no & "/" & total & ") Book=" & bookName & " Macro=" & macroName & " Error" )
			run = False
		End If
	End Property
End Class

Class ExcelRunnableCollection
	Private runnableArray
	Private count
	
	Public Sub Class_Initialize()
		Set runnableArray = CreateObject("System.Collections.ArrayList")
		count = 0
	End Sub

	Public Property Get add(m)
		count = count + 1
		Dim excelRunnable
		Set excelRunnable = new ExcelRunnable
		excelRunnable.setProperties m, count
		runnableArray.Add excelRunnable
		Set excelRunnable = Nothing
	End Property

	Public Sub execute(currentPath, bookName)
		Dim excel
		Set excel = CreateObject("Excel.Application")
		excel.Visible = False
		
		excel.Workbooks.Open currentPath & "\" & bookName
		
		isRun = True

		WScript.Echo ""
		WScript.Echo "Excel Runner Start"

		For Each runnable In runnableArray
			If isRun = True And runnable.run(excel, count) = False Then
				isRun = False
			End If
		Next

		excel.WorkBooks.Close
		excel.Quit
		Set excel = Nothing
		Set runnableArray = Nothing

		WScript.Echo "Excel Runner Complete"
		WScript.Echo ""
	End Sub
End Class

Class AccessRunnable
	Private macroName
	Private no

	Public Property Get setProperties(m, n)
		macroName = m
		no = n
	End Property

	Public Property Get run(access, currentPath, total)
		WScript.Echo( "(" & no & "/" & total & ") Macro=" & macroName & " Start" )
		result = access.Application.Run(macroName)
		If result = 0 Then
			WScript.Echo( "(" & no & "/" & total & ") Macro=" & macroName & " Complete" )
			run = True
		Else
			WScript.Echo( "(" & no & "/" & total & ") Macro=" & macroName & " Erro" )
			run = False
		End If
	End Property
End Class

Class AccessRunnableCollection
	Private runnableArray
	Private count

	Public Sub Class_Initialize()
		Set runnableArray = CreateObject("System.Collections.ArrayList")
		count = 0
	End Sub

	Public Property Get add(m)
		count = count + 1
		Dim accessRunnable
		Set accessRunnable = new AccessRunnable
		accessRunnable.setProperties m, count
		runnableArray.Add accessRunnable
		Set accessRunnable = Nothing
	End Property

	Public Sub execute(currentPath, dbName)
		Dim access
		Set access = CreateObject("Access.Application")
		access.Visible = False
		access.OpenCurrentDatabase currentPath & "\" & dbName

		isRun = True

		WScript.Echo ""
		WScript.Echo "Access Runner Start"

		For Each runnable In runnableArray
			If isRun = True And runnable.run(access, currentPath, count) = False Then
				isRun = False
			End If
		Next

		access.CloseCurrentDatabase
		access.Quit
		Set access = Nothing
		Set runnableArray = Nothing

		WScript.Echo "Access Runner Complete"
		WScript.Echo ""
	End Sub
End Class

