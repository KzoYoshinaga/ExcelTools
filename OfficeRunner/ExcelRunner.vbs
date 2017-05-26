Class ExcelRunnable
	Private bookName
	Private macroName
	Private no

	Public Property Get setProperties(b, m, n)
		bookName = b
		macroName = m
		no = n
	End Property

	Public Property Get run(excel, currentPath, total)
		excel.Workbooks.Open currentPath & "\" & bookName

		WScript.Echo( "(" & no & "/" & total & ") Book=" & bookName & " Macro=" & macroName & " Start" )

		result = excel.Application.Run(macroName)
		excel.Workbooks.Close
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

	Public Property Get add(b, m)
		count = count + 1
		Dim excelRunnable
		Set excelRunnable = new ExcelRunnable
		excelRunnable.setProperties b, m, count
		runnableArray.Add excelRunnable
		Set excelRunnable = Nothing
	End Property

	Public Sub execute(currentPath)
		Dim excel
		Set excel = CreateObject("Excel.Application")
		excel.Visible = False

		isRun = True

		WScript.Echo ""
		WScript.Echo "Excel Runner Start"

		For Each runnable In runnableArray
			If isRun = True And runnable.run(excel, currentPath, count) = False Then
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