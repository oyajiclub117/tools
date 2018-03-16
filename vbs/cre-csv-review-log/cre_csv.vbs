Option Explicit

'//**********************************************************
'//* Global var definitions
'//**********************************************************
'-- var Object
Dim g_objFso
Dim g_objCsvFile
Dim g_objWord

'-- varLong
Dim g_lngIdData

'//**********************************************************
'//* @procedure proc_document
'//**********************************************************
Private Sub proc_document( _
	byref p_objDoc _
	)
	'-- var Object
	Dim objDoc

	p_objDoc.SaveAs p_objDoc.Path & "\" & g_objFso.GetBaseName(p_objDoc.Name), 2

End Sub

'//**********************************************************
'//* @procedure main
'//**********************************************************
Private Sub main
	Dim objDoc(3)
	Dim objArg
	Dim objWshShell
	
	'-- var String
	Dim strDesktop
	Dim strPrefix

	Set g_objFso = CreateObject("Scripting.FileSystemObject")
	Set g_objWord = CreateObject("Word.Application")
	set objWshShell = WScript.CreateObject("WScript.Shell")
    strDesktop = objWshShell.SpecialFolders("Desktop")
	g_objWord.Visible = True
	Set g_objCsvFile = g_objFso.CreateTextFile(strDesktop & "\bus_tab_columns.csv")
	

	g_lngIdData = 0

	For Each objArg In WScript.Arguments
		Set objDoc(0) = g_objWord.Documents.Open(objArg,,True)
		
'		g_objWord.Application.StatusBar = "bookname=" & objDoc(0).Name
		
		proc_document objDoc(0)
		objDoc(0).Close False
	Next

	g_objWord.Application.StatusBar = ""

	Msgbox "create tables successfully", vbInformation, "èàóùäÆóπ"

	g_objCsvFile.Close

	g_objWord.Quit
End Sub
