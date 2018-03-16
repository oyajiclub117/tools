Option Explicit

'-- enum Excel
Const xlDown = -4121
Const xlUp = -4162

'-- enum FileSystemObject
Const ForWriting = 2
Const ForAppending = 8

'//**********************************************************
'//* Global var definitions
'//**********************************************************
'-- var Object
Dim g_objFso
Dim g_objCsvFile
Dim g_objExcel

'-- varLong
Dim g_lngIdData

'//**********************************************************
'//* @procedure proc_workbook
'//**********************************************************
Private Sub proc_workbook( _
	byref p_objWkb _
	)
	'-- var Object
	Dim objWks
	
	For Each objWks In p_objWkb.Worksheets
		g_objExcel.Application.StatusBar = "start sheet name=" & objWks.Name 
		proc_worksheet objWks
		g_objExcel.Application.StatusBar = "ended sheet name=" & objWks.Name 
	Next
End Sub

'//**********************************************************
'//* @procedure proc_worksheet
'//**********************************************************
Private Sub proc_worksheet( _
	byref p_objWks _
	)
	'-- var Object
	
	'-- var Long
	
	'-- var String
	Dim strCsvData
	
	g_lngIdData = g_lngIdData + 1
	
	strCsvData = _
		g_lngIdData & "," & _
		p_objWks.Name
		
	createCsvFile_Data strCsvData

End Sub

'//**********************************************************
'//* @procedure createCsvFile_Head
'//**********************************************************
Private Sub createCsvFile_Head
	'-- var String
	Dim strCsvHead
	
	strCsvHead = _
		"id" & "," & _
		"name"
	
	g_objCsvFile.WriteLine strCsvHead
End Sub

'//**********************************************************
'//* @procedure createCsvFile_Data
'//**********************************************************
Private Sub createCsvFile_Data( _
	byval p_strData _
	)
	g_objCsvFile.WriteLine p_strData
End Sub

'//**********************************************************
'//* @procedure main
'//**********************************************************
Private Sub main( _
	byval p_strFileName _
	)
	Dim objWkb(3)
	Dim objArg
	Dim objWshShell
	
	'-- var String
	Dim strDesktop
	Dim strPath

	Set g_objFso = CreateObject("Scripting.FileSystemObject")
	Set g_objExcel = CreateObject("Excel.Application")
	set objWshShell = WScript.CreateObject("WScript.Shell")
    strDesktop = objWshShell.SpecialFolders("Desktop")
	g_objExcel.Visible = True
	strPath = strDesktop & "\" & p_strFileName
'-	If g_objFso.FileExists(strPath) Then
'-		Set g_objCsvFile = g_objFso.OpenTextFile(strPath, ForAppending)
'-	Else
		Set g_objCsvFile = g_objFso.CreateTextFile(strPath)
		createCsvFile_Head
'-	End If
	

	g_lngIdData = 0

	For Each objArg In WScript.Arguments
		Set objWkb(0) = g_objExcel.Workbooks.Open(objArg,,True)
		
		g_objExcel.Application.StatusBar = "bookname=" & objWkb(0).Name
		
		proc_workbook objWkb(0)
		objWkb(0).Close False
	Next

	g_objExcel.Application.StatusBar = ""

	Msgbox "script execute successfully", vbInformation, "èàóùäÆóπ"

	g_objCsvFile.Close

	g_objExcel.Quit
End Sub
