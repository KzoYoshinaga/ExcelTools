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

	Publec Sub Class_Initialize()
		Set runnableArray = CreateObject("System.Collection.ArrayList")
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